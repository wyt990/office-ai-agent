## ADDED Requirements

### Requirement: Embedding availability check
The system SHALL provide a static method `EmbeddingService.IsEmbeddingAvailable()` that returns `Boolean`, indicating whether the current API configuration supports embedding generation. The check SHALL verify:
1. `ConfigSettings.EmbeddingModel` is not empty, OR `ConfigSettings.ApiUrl` can be mapped to a known default embedding model
2. `ConfigSettings.ApiUrl` is not in the known-unsupported list (e.g., providers that only offer chat completions without embedding endpoints)
3. `ConfigSettings.ApiKey` is not empty

#### Scenario: Embedding model explicitly configured
- **WHEN** `ConfigSettings.EmbeddingModel` is set to a non-empty value AND `ConfigSettings.ApiKey` is non-empty
- **THEN** `IsEmbeddingAvailable()` SHALL return `True`

#### Scenario: API URL maps to known default model
- **WHEN** `ConfigSettings.EmbeddingModel` is empty AND `ConfigSettings.ApiUrl` contains a recognized provider (e.g., "siliconflow", "dashscope", "openai")
- **THEN** `IsEmbeddingAvailable()` SHALL return `True`

#### Scenario: API provider does not support embedding
- **WHEN** `ConfigSettings.ApiUrl` matches a known-unsupported provider (e.g., contains "deepseek") AND `ConfigSettings.EmbeddingModel` is empty
- **THEN** `IsEmbeddingAvailable()` SHALL return `False`

#### Scenario: API key missing
- **WHEN** `ConfigSettings.ApiKey` is empty or whitespace
- **THEN** `IsEmbeddingAvailable()` SHALL return `False`

### Requirement: Graceful degradation on embedding generation
The system SHALL NOT throw exceptions when embedding generation fails or is unavailable. `EmbeddingService.GetEmbeddingAsync()` SHALL call `IsEmbeddingAvailable()` at entry and return `Nothing` immediately when unavailable, without issuing an HTTP request.

#### Scenario: Embedding unavailable at generation time
- **WHEN** `IsEmbeddingAvailable()` returns `False` AND `GetEmbeddingAsync()` is called
- **THEN** the method SHALL return `Nothing` without making any HTTP request AND SHALL log a debug message indicating the reason

#### Scenario: HTTP request fails at runtime
- **WHEN** `IsEmbeddingAvailable()` returns `True` but the HTTP embedding request fails (non-200 status or network error)
- **THEN** the method SHALL return `Nothing` AND SHALL cache the failure for 30 minutes to avoid repeated failed requests

#### Scenario: Cached failure within cooldown period
- **WHEN** a previous embedding request failed within the last 30 minutes
- **THEN** `GetEmbeddingAsync()` SHALL return `Nothing` immediately without issuing a new HTTP request

### Requirement: Memory saving without embedding
The system SHALL save atomic memories to the database even when embedding generation fails or is unavailable. The `embedding` column SHALL be `NULL` in this case, and the record SHALL remain searchable via keyword (LIKE) fallback.

#### Scenario: Save memory when embedding unavailable
- **WHEN** `MemoryService.SaveAtomicMemoryAsync()` is called AND embedding generation returns `Nothing`
- **THEN** the memory record SHALL be inserted into `atomic_memory` with `embedding = NULL` AND a debug log SHALL indicate the record was saved without vector

#### Scenario: Save memory when embedding succeeds
- **WHEN** `MemoryService.SaveAtomicMemoryAsync()` is called AND embedding generation returns a valid vector
- **THEN** the memory record SHALL be inserted with the serialized embedding in the `embedding` column

### Requirement: RAG query without embedding
The system SHALL perform RAG retrieval even when query embedding generation fails. `MemoryService.GetRelevantMemories()` SHALL skip vector generation when `IsEmbeddingAvailable()` returns `False` and proceed directly to keyword-based (LIKE) fallback search.

#### Scenario: RAG search when embedding unavailable
- **WHEN** `GetRelevantMemories()` is called AND `IsEmbeddingAvailable()` returns `False`
- **THEN** the system SHALL skip the embedding generation step AND SHALL use LIKE-based keyword search directly

#### Scenario: RAG search when embedding available but generation fails
- **WHEN** `GetRelevantMemories()` is called AND `IsEmbeddingAvailable()` returns `True` but `GetEmbeddingAsync()` returns `Nothing` (runtime failure)
- **THEN** the system SHALL fall back to LIKE-based keyword search

### Requirement: Configuration warning for missing embedding model
When saving API configuration in `ConfigApiForm`, if no embedding model is selected and the API provider is not in the known-supported list, the system SHALL display a non-blocking warning to the user.

#### Scenario: Save config without embedding model on unknown provider
- **WHEN** user saves API configuration AND `EmbeddingModel` is empty AND the API URL does not match any known embedding-supporting provider
- **THEN** the system SHALL display a warning message indicating that Memory/RAG features will use keyword search only

#### Scenario: Save config with embedding model selected
- **WHEN** user saves API configuration AND `EmbeddingModel` is set to a non-empty value
- **THEN** the system SHALL NOT display any embedding-related warning
