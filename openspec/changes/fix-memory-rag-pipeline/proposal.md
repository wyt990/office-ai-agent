## Why

当前 Memory/RAG 管线存在 5 个相互关联的缺陷，导致 AI 聊天体验严重受损：向量模型缺失时直接崩溃、RAG 召回内容与当前对话不相关、新对话无法实时入库、记忆无分层且 assistant 回复丢失、Agent 模式的系统提示词被重复全量发送浪费 token。这些问题使得"记忆增强对话"功能基本不可用，需要系统性修复。

## What Changes

- **向量模型守卫**：在 `EmbeddingService.GetEmbeddingAsync()` 和 `MemoryService.SaveAtomicMemoryAsync()` 中增加 embedding model 可用性校验，模型为空或不可用时优雅降级（跳过向量存储、使用关键词回退检索），不再抛出异常
- **RAG 召回质量提升**：在 `MemoryRepository.GetRelevantMemories()` 中引入相似度阈值过滤（丢弃低于阈值的结果）、限制返回条数、按 app_type 过滤，增加时间衰减因子让近期记忆权重更高
- **实时入库修复**：在 `BaseChatControl` 的流式响应完成回调中，将完整的 user 消息（即发送给大模型的实际内容）和 assistant 回复同时存入向量库，而非仅存用户原始输入
- **长期/短期记忆分层**：在 `MemoryRepository` 的 atomic_memories 表中增加 `memory_type` 字段（`short_term` / `long_term`），短期记忆为当前会话的对话轮次（含 role=assistant），长期记忆为跨会话的摘要或用户显式标记的内容；`ChatContextBuilder` 分层注入：短期记忆作为近期上下文窗口，长期记忆通过 RAG 召回
- **系统提示词单次注入**：重构 `RalphAgentController` 中规划提示词的发送方式，将 `PLANNING_PROMPT_*` 作为 `role=system` 消息仅在会话首次规划时发送一次，后续轮次仅发送增量的 user/assistant 消息；`ChatContextBuilder.BuildMessages()` 确保 system 消息不重复

## Capabilities

### New Capabilities
- `embedding-guard`: 向量模型可用性校验与优雅降级机制，覆盖配置校验、运行时守卫、降级策略
- `memory-tiering`: 长期/短期记忆分层存储与检索，覆盖数据模型、存储策略、分层注入上下文
- `rag-recall-quality`: RAG 召回质量优化，覆盖相似度阈值、时间衰减、条数限制、app_type 过滤

### Modified Capabilities
_(无已有 specs 需要修改)_

## Impact

- **数据库变更**：`atomic_memories` 表需 `ALTER TABLE ADD COLUMN memory_type`，需迁移脚本
- **受影响文件**：
  - `ShareRibbon\Services\EmbeddingService.vb` — 增加模型校验
  - `ShareRibbon\Controls\Services\MemoryService.vb` — 入库逻辑修改、分层存储
  - `ShareRibbon\Storage\MemoryRepository.vb` — 检索质量优化、新字段支持
  - `ShareRibbon\Controls\Services\ChatContextBuilder.vb` — 分层注入、system 去重
  - `ShareRibbon\Controls\BaseChatControl.vb` — 实时入库时机和内容修正
  - `ShareRibbon\Loop\RalphAgentController.vb` — 系统提示词单次发送
  - `ShareRibbon\Config\ConfigApiForm.vb` — 向量模型配置校验提示
- **兼容性**：数据库迁移向后兼容（新字段有默认值），无 API 接口变更，无 **BREAKING** 变更
