## 1. 数据库迁移

- [x] 1.1 在 `OfficeAiDatabase.RunVersionedMigrations()` 中添加版本 6 迁移：`ALTER TABLE atomic_memory ADD COLUMN memory_type TEXT DEFAULT 'short_term'`，更新 schema_version 到 6
- [x] 1.2 在 `AtomicMemoryRecord` 实体类中添加 `MemoryType As String` 属性
- [x] 1.3 更新 `MemoryRepository` 中所有 SELECT 查询以包含 `memory_type` 列，并在 reader 映射中填充 `MemoryType` 属性

## 2. Embedding 守卫

- [x] 2.1 在 `EmbeddingService` 中添加 `IsEmbeddingAvailable() As Boolean` 静态方法，检查 EmbeddingModel + ApiUrl + ApiKey 的可用性，维护已知不支持 embedding 的 provider 列表（如 DeepSeek）
- [x] 2.2 在 `EmbeddingService` 中添加失败缓存机制：私有静态字段 `_lastFailureTime As DateTime?`，在 HTTP 请求失败时记录时间，30 分钟内不再重试
- [x] 2.3 修改 `EmbeddingService.GetEmbeddingAsync()` 入口，调用 `IsEmbeddingAvailable()` 和检查失败缓存，不可用时直接返回 `Nothing`
- [x] 2.4 修改 `MemoryService.GetRelevantMemories()`，在调用 `GetEmbeddingAsync` 前先检查 `IsEmbeddingAvailable()`，不可用时跳过向量生成直接走 LIKE 回退
- [x] 2.5 在 `ConfigApiForm.vb` 保存配置逻辑中，当 EmbeddingModel 为空且 API 不在已知支持列表时显示非阻塞警告

## 3. RAG 召回质量优化

- [x] 3.1 在 `MemoryConfig` 中添加 `RagSimilarityThreshold`（Single，默认 0.3，范围 0.0–1.0）和 `RagTimeDecayRate`（Single，默认 0.01，范围 0.0–1.0）两个配置属性，含 `MemoryConfigData` 对应字段
- [x] 3.2 修改 `MemoryRepository.GetRelevantMemories()` 签名，增加 `appType As String` 参数
- [x] 3.3 修改向量检索分支 SQL，添加 `AND (app_type = @app OR app_type IS NULL OR app_type = '')` 条件（当 appType 非空时）
- [x] 3.4 在向量检索分支中实现时间衰减计算：`finalScore = similarity * (1 / (1 + daysSinceCreation * decayRate))`
- [x] 3.5 在向量检索分支中实现相似度阈值过滤：丢弃 `finalScore < RagSimilarityThreshold` 的结果
- [x] 3.6 修改 LIKE 回退分支 SQL，添加相同的 `app_type` 过滤条件
- [x] 3.7 更新 `MemoryService.GetRelevantMemories()` 和 `MemoryService.SearchMemories()` 调用链，传递 `appType` 参数
- [x] 3.8 更新 `ChatContextBuilder.BuildMessages()` 和 `RalphAgentController.StartAgent()` 中对 `GetRelevantMemories` 的调用，传入当前 `appType`

## 4. 实时入库与分条存储

- [x] 4.1 在 `MemoryService` 中新增 `SaveConversationTurnAsync(userContent, assistantContent, sessionId, appType)` 方法，分别存储 user 和 assistant 两条记忆记录
- [x] 4.2 实现改进的去重逻辑：基于 `session_id + 内容前 50 字符` 查询已有记录，匹配则跳过
- [x] 4.3 修改 `MemoryRepository.InsertAtomicMemory()` 签名，增加 `memoryType As String` 参数（默认 `"short_term"`）
- [x] 4.4 修改 `BaseChatControl.SendHttpRequestStream()` finally 块（约行 4085），将 `MemoryService.SaveAtomicMemoryAsync()` 调用替换为 `SaveConversationTurnAsync(result, answer.content, sessionId, appType)`，其中 `result` 是经 StripQuestion 处理后的实际发送内容
- [x] 4.5 修改 Agent 模式的记忆保存（约行 2321），同样使用 `SaveConversationTurnAsync`

## 5. 记忆分层注入

- [x] 5.1 修改 `MemoryRepository.GetRelevantMemories()` 的向量检索和 LIKE 回退分支 SQL，增加 `AND memory_type = 'long_term'` 过滤条件，避免短期记忆与会话窗口重复
- [x] 5.2 确认 `ChatContextBuilder.BuildMessages()` 层 [5] 的 sessionMessages 已包含 `role=assistant` 消息（当前实现已正确，需验证）
- [x] 5.3 确认 `BaseChatControl` 在流式响应完成后将 assistant 回复添加到 `systemHistoryMessageData`（当前行 4080 已实现，需验证无遗漏）

## 6. Agent 系统提示词拆分

- [x] 6.1 将 `PLANNING_PROMPT_EXCEL` 拆分为 `PLANNING_SYSTEM_PROMPT_EXCEL`（固定指令：角色定义、JSON 格式要求、command 文档）和 `PLANNING_USER_TEMPLATE_EXCEL`（动态内容占位符：当前内容、用户需求、意图、RAG、历史）
- [x] 6.2 对 `PLANNING_PROMPT_WORD` 和 `PLANNING_PROMPT_POWERPOINT` 执行相同拆分
- [x] 6.3 修改 `GetPlanningPrompt()` 返回 Tuple(Of String, String) 分别返回 system 和 user 模板
- [x] 6.4 修改 `StartAgent()` 中的调用：`SendAIRequest(userPrompt, systemPrompt)` 替换原 `SendAIRequest(prompt, "")`
- [x] 6.5 将 `STEP_EXECUTION_PROMPT` 同样拆分为 system 部分（格式要求、command 文档）和 user 部分（步骤详情），修改 `ExecuteNextStep()` 中的调用

## 7. 编译与验证

- [x] 7.1 编译 ShareRibbon 项目（`msbuild AiHelper.sln`），修复编译错误
- [x] 7.2 检查所有 `MemoryRepository.GetRelevantMemories()` 调用点是否已更新新签名（含 appType 参数）
- [x] 7.3 检查所有 `MemoryRepository.InsertAtomicMemory()` 调用点是否已适配新签名（含 memoryType 参数）
