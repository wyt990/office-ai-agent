## ADDED Requirements

### Requirement: Memory type field in database
The `atomic_memory` table SHALL have a `memory_type TEXT DEFAULT 'short_term'` column, added via database migration version 6. Valid values are `'short_term'` and `'long_term'`.

#### Scenario: Database upgrade from version 5
- **WHEN** the application starts with database schema version 5
- **THEN** the migration SHALL execute `ALTER TABLE atomic_memory ADD COLUMN memory_type TEXT DEFAULT 'short_term'` AND update schema version to 6

#### Scenario: Existing records retain default
- **WHEN** the migration completes on a database with existing atomic_memory records
- **THEN** all existing records SHALL have `memory_type = 'short_term'`

### Requirement: Conversation turn storage with role separation
The system SHALL store each conversation turn as two separate atomic memory records: one for the user message (`role=user`) and one for the assistant reply (`role=assistant`). A new method `MemoryService.SaveConversationTurnAsync(userContent, assistantContent, sessionId, appType)` SHALL handle this.

#### Scenario: Normal conversation turn saved
- **WHEN** a chat conversation completes with a user message and an assistant reply
- **THEN** the system SHALL create two `atomic_memory` records: one with the user's message content and one with the assistant's reply content, both with `memory_type = 'short_term'` and the same `session_id`

#### Scenario: User content is the processed message
- **WHEN** `SaveConversationTurnAsync` is called from `BaseChatControl`
- **THEN** the `userContent` parameter SHALL be the processed message (after `StripQuestion`) that was actually sent to the LLM, NOT the raw user input before processing

#### Scenario: Assistant content is the full reply
- **WHEN** `SaveConversationTurnAsync` is called from `BaseChatControl`
- **THEN** the `assistantContent` parameter SHALL be the complete assistant reply (`allMarkdownBuffer` content), not a truncated version

### Requirement: Deduplication by session and role
The system SHALL prevent duplicate memory records for the same conversation turn. Deduplication SHALL be based on matching `session_id` + content prefix (first 50 characters), checked separately for user and assistant records.

#### Scenario: Same turn not saved twice
- **WHEN** `SaveConversationTurnAsync` is called twice with the same `sessionId` and content that shares the same first 50 characters
- **THEN** the second call SHALL skip insertion for the duplicate record(s)

#### Scenario: Different turns in same session
- **WHEN** two different conversation turns occur in the same session with different content
- **THEN** both turns SHALL be saved as separate records

### Requirement: Short-term memory via session window
Short-term memories (current session conversation turns) SHALL be injected into the LLM context via the existing session messages sliding window (layer [5] in `ChatContextBuilder`), NOT via RAG retrieval. The assistant replies in `systemHistoryMessageData` SHALL be included in the sliding window.

#### Scenario: Assistant replies in sliding window
- **WHEN** `ChatContextBuilder.BuildMessages()` processes `sessionMessages`
- **THEN** messages with `role = "assistant"` SHALL be included in the output alongside `role = "user"` messages

#### Scenario: Short-term memories excluded from RAG
- **WHEN** `MemoryService.GetRelevantMemories()` performs RAG retrieval
- **THEN** the SQL query SHALL include `AND memory_type = 'long_term'` to exclude short-term memories from RAG results, avoiding duplication with the sliding window

### Requirement: Long-term memory via RAG recall
Long-term memories SHALL be retrieved via vector similarity search (or LIKE fallback) in `MemoryService.GetRelevantMemories()` and injected into the LLM context at layer [3][4] of `ChatContextBuilder`.

#### Scenario: Only long-term memories in RAG results
- **WHEN** the database contains both `short_term` and `long_term` memories AND a RAG query is executed
- **THEN** only memories with `memory_type = 'long_term'` SHALL appear in the RAG results

#### Scenario: User-promoted long-term memory
- **WHEN** a user explicitly marks a memory as important (via the management UI)
- **THEN** the memory's `memory_type` SHALL be updated to `'long_term'`

### Requirement: AtomicMemoryRecord includes memory type
The `AtomicMemoryRecord` entity class SHALL include a `MemoryType As String` property, populated from the `memory_type` column when reading from the database.

#### Scenario: Record loaded with memory type
- **WHEN** an `AtomicMemoryRecord` is read from the database
- **THEN** the `MemoryType` property SHALL contain the value from the `memory_type` column (either `'short_term'` or `'long_term'`)

### Requirement: Agent system prompt sent as role=system
`RalphAgentController` SHALL send the planning prompt's fixed instructions (format requirements, command documentation) as `role=system` message, and the dynamic content (user request, current content, intent, RAG memories, history) as `role=user` message.

#### Scenario: Planning request message structure
- **WHEN** `RalphAgentController.StartAgent()` calls `SendAIRequest`
- **THEN** the second parameter (systemPrompt) SHALL contain the fixed planning instructions AND the first parameter (userPrompt) SHALL contain only the dynamic user-specific content

#### Scenario: Step execution prompt structure
- **WHEN** `RalphAgentController.ExecuteNextStep()` calls `SendAIRequest` for code generation
- **THEN** the second parameter (systemPrompt) SHALL contain the fixed step execution instructions AND the first parameter SHALL contain the step-specific details

### Requirement: System prompt deduplication in ChatContextBuilder
`ChatContextBuilder.BuildMessages()` SHALL ensure that the final message list contains at most one `role=system` message. All system-level content (base prompt, scenario instructions, skills, memory blocks) SHALL be merged into a single system message.

#### Scenario: Single system message in output
- **WHEN** `BuildMessages()` is called with a non-empty `baseSystemPrompt` AND scenario instructions exist AND memory is enabled
- **THEN** the output message list SHALL contain exactly one message with `role = "system"` at index 0, with all system content merged

#### Scenario: No system content
- **WHEN** `BuildMessages()` is called with an empty `baseSystemPrompt` AND no scenario instructions AND memory disabled
- **THEN** the output message list SHALL NOT contain any `role = "system"` message
