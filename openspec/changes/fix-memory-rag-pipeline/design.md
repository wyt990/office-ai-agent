## Context

当前 Memory/RAG 管线由以下组件协作：

- **EmbeddingService** — 调用 OpenAI 格式的 embedding API 生成向量，通过 `ConfigSettings.EmbeddingModel` 或 `GetDefaultEmbeddingModel()` 确定模型名
- **MemoryService** — 封装 RAG 检索、原子记忆写入，是 `BaseChatControl` 和 `RalphAgentController` 的统一入口
- **MemoryRepository** — SQLite CRUD，`GetRelevantMemories()` 先全量加载再内存排序
- **ChatContextBuilder** — 分层组装 `[0]system → [1]场景+Skills → [3][4]记忆 → [5]会话窗口 → [6]user`，所有记忆块追加到 system 消息尾部
- **BaseChatControl.CreateRequestBody()** — 调用 `ChatContextBuilder.BuildMessages()` 构建最终请求消息数组
- **RalphAgentController** — Agent 模式下，将规划提示词作为 user 消息发送（`SendAIRequest(prompt, "")`），system 为空字符串

数据库 schema 当前版本为 5，`atomic_memory` 表有 `id, timestamp, content, tags, session_id, create_time, app_type, embedding` 字段。

## Goals / Non-Goals

**Goals:**

- G1: embedding model 为空或 API 不支持时不崩溃，优雅降级到关键词检索
- G2: RAG 召回结果与当前对话高度相关（相似度阈值 + 时间衰减 + app_type 过滤）
- G3: 每轮对话完成后将完整的 user 发送内容和 assistant 回复分别入库（含 embedding），可立即被后续 RAG 检索到
- G4: 区分短期记忆（当前会话对话轮次）和长期记忆（跨会话 RAG），assistant 回复以 `role=assistant` 的语义存入
- G5: Agent 规划提示词作为 `role=system` 仅首轮发送，后续轮次仅传增量 user/assistant

**Non-Goals:**

- 不引入外部向量数据库（如 Milvus/Qdrant），继续使用 SQLite
- 不重构 ConfigApiForm.vb 的整体 UI，仅增加 embedding model 缺失提示
- 不改变 `ChatContextBuilder.BuildMessages()` 的 7 层分层结构
- 不实现自动长期记忆摘要（由大模型总结跨会话记忆），留作后续迭代

## Decisions

### D1: Embedding 守卫策略 — "检测 + 标记 + 降级"

**选择**：在 `EmbeddingService.GetEmbeddingAsync()` 入口处增加前置校验，若 embedding model 为空且默认推断也不可靠，返回 `Nothing` 而非发起 HTTP 请求。同时在 `EmbeddingService` 中增加 `IsEmbeddingAvailable()` 静态方法供其他组件预判。

**关键改动**：
- `EmbeddingService.IsEmbeddingAvailable()` — 检查 `ConfigSettings.EmbeddingModel` 和 `ConfigSettings.ApiUrl` 是否支持 embedding（DeepSeek 等已知不支持的排除）
- `EmbeddingService.GetEmbeddingAsync()` — 入口处调用 `IsEmbeddingAvailable()`，不可用时直接返回 `Nothing`
- `MemoryService.SaveAtomicMemoryAsync()` — embedding 生成失败时仍保存记忆（`embedding=NULL`），确保关键词检索可用
- `MemoryService.GetRelevantMemories()` — embedding 不可用时跳过向量生成，直接走 LIKE 回退
- `ConfigApiForm.vb` — 保存配置时若未选择 embedding model 且 API 不在已知支持列表中，显示警告提示

**备选方案**（已否决）：在 `ConfigApiForm` 保存时阻止——过于激进，部分用户可能不需要 Memory 功能。

### D2: RAG 召回质量 — "阈值 + 衰减 + 过滤"

**选择**：在 `MemoryRepository.GetRelevantMemories()` 的向量检索分支中引入三个过滤维度：

1. **相似度阈值**：余弦相似度低于 0.3 的结果丢弃（可通过 `MemoryConfig.RagSimilarityThreshold` 配置）
2. **时间衰减因子**：`finalScore = similarity * timeDecay`，其中 `timeDecay = 1 / (1 + daysSinceCreation * decayRate)`，`decayRate` 默认 0.01（约 100 天半衰期）
3. **app_type 过滤**：在 SQL WHERE 中增加 `AND (app_type = @app OR app_type IS NULL OR app_type = '')`，仅检索当前宿主相关记忆

**关键改动**：
- `MemoryConfig` 新增 `RagSimilarityThreshold`（默认 0.3）和 `RagTimeDecayRate`（默认 0.01）
- `MemoryRepository.GetRelevantMemories()` 新增 `appType` 参数，SQL 增加 app_type 过滤
- 向量检索分支：计算 `finalScore` 并过滤低于阈值的结果
- LIKE 回退分支：同样增加 app_type 过滤

**备选方案**（已否决）：使用 SQLite FTS5 全文索引——虽然更精确，但需要额外的 FTS 表维护和中文分词器配置，复杂度过高。

### D3: 实时入库 — "分条存储 + 完整内容"

**选择**：修改 `BaseChatControl` 的入库逻辑，将 user 消息和 assistant 回复分别存储为独立的原子记忆记录，而非截断合并。

**关键改动**：
- `MemoryService.SaveAtomicMemoryAsync()` 重构签名，接受 `role` 参数（`"user"` 或 `"assistant"`），分别存储
- `MemoryService` 新增 `SaveConversationTurnAsync(userContent, assistantContent, sessionId, appType)` 方法，一次调用存两条记忆
- `BaseChatControl` 行 4085：调用 `SaveConversationTurnAsync()`，`userContent` 使用 `result`（经 StripQuestion 处理后实际发给大模型的内容），`assistantContent` 使用 `answer.content`
- 去重逻辑改进：基于 `session_id + role + 内容前 50 字符` 做去重，避免同一轮对话重复入库

**备选方案**（已否决）：只存一条合并记录——导致 RAG 检索时无法区分用户问题和 AI 回答，降低召回精度。

### D4: 记忆分层 — "字段标记 + 注入分离"

**选择**：在 `atomic_memory` 表增加 `memory_type TEXT DEFAULT 'short_term'` 字段（数据库迁移版本 6），通过字段区分而非分表。

**分层规则**：
- `short_term`：当前会话的每轮对话（user + assistant），随会话结束可自动过期或晋升
- `long_term`：用户显式收藏的记忆、跨会话被多次 RAG 命中的高频记忆（后续迭代实现自动晋升）

**上下文注入分离**（`ChatContextBuilder`）：
- 短期记忆：通过现有的 `sessionMessages`（滚动窗口层 [5]）注入，不走 RAG
- 长期记忆：通过 `MemoryService.GetRelevantMemories()` 走 RAG 召回，注入层 [3][4]
- `GetRelevantMemories()` 的 SQL 查询增加 `AND memory_type = 'long_term'` 过滤，避免短期记忆与滚动窗口重复

**数据库迁移**：版本 6，`ALTER TABLE atomic_memory ADD COLUMN memory_type TEXT DEFAULT 'short_term'`

**备选方案**（已否决）：分两张表（`short_term_memory` / `long_term_memory`）——增加代码复杂度，且迁移成本高。

### D5: 系统提示词单次发送 — "拆分 prompt 为 system + user"

**选择**：重构 `RalphAgentController.StartAgent()` 中规划请求的构建方式。

**当前问题**：
- `SendAIRequest(prompt, "")` 的第二个参数是空字符串，意味着 system prompt 为空
- `prompt` 变量包含了完整的规划指令（~100 行模板 + 用户内容 + 意图 + RAG 记忆），全部作为 user 消息发送
- 每次规划都发送完整模板，浪费 token

**重构方案**：
- 将 `PLANNING_PROMPT_EXCEL/WORD/POWERPOINT` 拆分为两部分：
  - **系统部分**（固定指令、格式要求、command 文档）→ 作为 `systemPrompt` 参数传递
  - **用户部分**（当前内容、用户需求、意图、RAG 记忆、历史对话）→ 作为 `userPrompt` 参数传递
- `SendAIRequest` 调用改为 `SendAIRequest(userPrompt, systemPrompt)`
- 对于步骤执行（`STEP_EXECUTION_PROMPT`），同样拆分 system 与 user

**备选方案**（已否决）：在 `RalphAgentController` 内部维护会话历史并跟踪 system 是否已发——增加状态管理复杂度，且 `SendAIRequest` 已有 `BaseChatControl` 的 system 处理逻辑。

## Risks / Trade-offs

- **[R1] 相似度阈值可能过滤掉有用记忆** → 提供 `MemoryConfig.RagSimilarityThreshold` 可配置，默认 0.3 较保守；LIKE 回退确保零结果时仍有候选
- **[R2] 时间衰减可能淘汰旧但重要的记忆** → `decayRate=0.01` 意味着 100 天后仍保留约 50% 权重，长期记忆不受衰减影响（后续迭代可加 importance 字段）
- **[R3] 分条存储增加数据库记录量** → 单条对话从 1 条变 2 条，但 `AtomicContentMaxLength` 已限制内容长度，SQLite 可轻松承受
- **[R4] 数据库迁移版本 6** → `ALTER TABLE ADD COLUMN` 向后兼容，默认值 `'short_term'` 确保旧记录正常工作
- **[R5] DeepSeek 等不支持 embedding 的 API** → `IsEmbeddingAvailable()` 维护已知不支持列表，可能不全 → 兜底：HTTP 调用失败后缓存结果 30 分钟不再重试
