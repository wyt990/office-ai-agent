## ADDED Requirements

### Requirement: Similarity threshold filtering
The system SHALL discard RAG results whose cosine similarity score is below a configurable threshold. The threshold SHALL be stored in `MemoryConfig.RagSimilarityThreshold` with a default value of 0.3.

#### Scenario: Results above threshold
- **WHEN** vector retrieval returns memories with cosine similarity scores [0.85, 0.62, 0.45, 0.28, 0.15] AND the threshold is 0.3
- **THEN** the system SHALL return only the 3 memories with scores [0.85, 0.62, 0.45] AND discard the remaining 2

#### Scenario: All results below threshold
- **WHEN** vector retrieval returns memories all with cosine similarity below 0.3
- **THEN** the system SHALL return an empty list from the vector branch AND fall back to LIKE-based keyword search

#### Scenario: Threshold is configurable
- **WHEN** user sets `MemoryConfig.RagSimilarityThreshold` to 0.5
- **THEN** the system SHALL use 0.5 as the threshold for all subsequent RAG queries

### Requirement: Time decay scoring
The system SHALL apply a time decay factor to each memory's cosine similarity score, so that recently created memories rank higher than older ones with the same raw similarity. The formula SHALL be: `finalScore = similarity * (1 / (1 + daysSinceCreation * decayRate))` where `decayRate` is stored in `MemoryConfig.RagTimeDecayRate` with a default of 0.01.

#### Scenario: Recent memory ranked higher
- **WHEN** two memories have identical cosine similarity of 0.7 AND memory A was created 1 day ago AND memory B was created 100 days ago AND `decayRate` is 0.01
- **THEN** memory A's final score SHALL be approximately `0.7 * (1/1.01) ≈ 0.693` AND memory B's final score SHALL be approximately `0.7 * (1/2.0) = 0.35` AND memory A SHALL rank higher

#### Scenario: Decay rate is configurable
- **WHEN** user sets `MemoryConfig.RagTimeDecayRate` to 0.0 (no decay)
- **THEN** the system SHALL not apply time decay and ranking SHALL be based solely on cosine similarity

### Requirement: App type filtering
The system SHALL filter RAG results by the current Office application type (`appType`), returning only memories that were created in the same application context or have no application type set.

#### Scenario: Filter by Excel app type
- **WHEN** RAG query is made from an Excel add-in AND the database contains memories with `app_type` values "Excel", "Word", "", and NULL
- **THEN** the system SHALL return only memories where `app_type = 'Excel'` OR `app_type IS NULL` OR `app_type = ''`

#### Scenario: No app type provided
- **WHEN** RAG query is made without specifying `appType` (or `appType` is empty)
- **THEN** the system SHALL return memories regardless of their `app_type` value

### Requirement: App type filtering in LIKE fallback
The system SHALL apply the same `appType` filter in the LIKE-based keyword fallback path, ensuring consistent filtering regardless of whether vector or keyword search is used.

#### Scenario: LIKE search with app type
- **WHEN** embedding is unavailable AND LIKE-based search is performed from a Word add-in
- **THEN** the SQL query SHALL include `AND (app_type = 'Word' OR app_type IS NULL OR app_type = '')` in the WHERE clause

### Requirement: Combined scoring and sorting
After applying similarity threshold filtering and time decay, the system SHALL sort results by `finalScore` in descending order and return the top N results (where N is `MemoryConfig.RagTopN`).

#### Scenario: Top N results returned
- **WHEN** 10 memories pass the similarity threshold AND `RagTopN` is 5
- **THEN** the system SHALL return the 5 memories with the highest `finalScore` values

#### Scenario: Fewer results than TopN
- **WHEN** only 2 memories pass the similarity threshold AND `RagTopN` is 5
- **THEN** the system SHALL return all 2 passing memories

### Requirement: New configuration properties
`MemoryConfig` SHALL expose two new configurable properties: `RagSimilarityThreshold` (Single, default 0.3, range 0.0–1.0) and `RagTimeDecayRate` (Single, default 0.01, range 0.0–1.0). These SHALL be persisted in the memory configuration JSON file.

#### Scenario: Default values on fresh install
- **WHEN** no memory configuration file exists
- **THEN** `RagSimilarityThreshold` SHALL default to 0.3 AND `RagTimeDecayRate` SHALL default to 0.01

#### Scenario: Values clamped to valid range
- **WHEN** user sets `RagSimilarityThreshold` to 1.5
- **THEN** the system SHALL clamp the value to 1.0
