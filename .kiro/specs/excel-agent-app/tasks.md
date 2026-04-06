# Implementation Plan: Excel Agent App

## Overview

Build a full-stack AI agent application with a Rust backend (adk-rust v0.5.0) and React frontend (TypeScript, Vite, Tailwind CSS). The backend spawns the Excel MCP server as a child process, configures an LLM agent with MCP tools, and serves an HTTP API with SSE streaming plus the built frontend. The frontend provides a split-pane chat + spreadsheet preview UI that consumes SSE events in real time. Implementation proceeds bottom-up: backend config and MCP lifecycle first, then agent setup, custom routes, frontend scaffolding, hooks, components, and finally integration wiring.

Assumes the excel-mcp-server binary is already built and available at a configurable path.

## Tasks

- [ ] 1. Initialize backend project and implement configuration
  - [x] 1.1 Create Cargo project and configure dependencies
    - Initialize a new Rust binary crate in the project root (or a `backend/` directory)
    - Add dependencies to `Cargo.toml`: `adk-agent`, `adk-tool`, `adk-server`, `adk-session`, `adk-model`, `adk-runner` (all via path deps to `~/Developer/projects/adk-rust`), `axum`, `tokio` (features: `full`), `serde` (features: `derive`), `serde_json`, `tracing`, `tracing-subscriber`, `tower`, `tower-http` (features: `fs`, `cors`), `uuid` (features: `v4`), `anyhow`
    - Create module structure: `src/main.rs`, `src/config.rs`, `src/agent.rs`, `src/mcp.rs`, `src/routes/mod.rs`, `src/routes/upload.rs`, `src/routes/artifacts.rs`, `src/artifacts.rs`, `src/static_files.rs`
    - _Requirements: 1.1, 1.4_

  - [ ] 1.2 Implement AppConfig and ModelProvider (`src/config.rs`)
    - Define `AppConfig` struct with fields: `host`, `port`, `model_provider`, `mcp_server_path`, `artifact_dir`, `static_dir`, `environment`, `allowed_origins`
    - Define `ModelProvider` enum with `Gemini`, `OpenAi`, `Anthropic` variants each holding `api_key` and `model` with defaults (gemini-2.5-flash, gpt-4o, claude-sonnet-4-20250514)
    - Define `Environment` enum (`Development`, `Production`)
    - Implement `AppConfig::from_env()` that reads all environment variables, validates `MODEL_PROVIDER` is one of "gemini"/"openai"/"anthropic", checks the required API key is set, and returns descriptive errors naming the missing variable
    - Default `ENVIRONMENT` to Development with a logged warning when unset
    - _Requirements: 2.1, 2.2, 2.3, 2.4, 2.5, 17.1, 17.2, 17.3_

  - [ ]* 1.3 Write property tests for config parsing
    - **Property 1: Model provider parsing accepts only valid values**
    - **Validates: Requirements 2.1**

  - [ ]* 1.4 Write property tests for missing API key detection
    - **Property 2: Missing API key produces error naming the variable**
    - **Validates: Requirements 2.5**

- [ ] 2. Implement MCP server lifecycle and agent builder
  - [ ] 2.1 Implement MCP server spawn and monitoring (`src/mcp.rs`)
    - Implement `spawn_mcp_server(path)` that uses `TokioChildProcess` to spawn the Excel MCP server binary with stdin/stdout/stderr piped, connects the MCP client, and creates an `McpToolset`
    - Define `McpHealthStatus` enum: `Healthy`, `Degraded { reason }`, `Critical { reason }`
    - Implement `monitor_and_respawn()` background task that watches the child process, attempts up to 3 respawn attempts on crash, and updates a shared `Arc<RwLock<McpHealthStatus>>`
    - Log all spawn/crash/respawn events to stderr via `tracing`
    - _Requirements: 1.1, 1.2, 1.5, 8.3, 8.4, 8.5_

  - [ ] 2.2 Implement agent builder (`src/agent.rs`)
    - Implement `build_agent(config, toolset)` that creates the appropriate model (`GeminiModel`, `OpenAiModel`, or `AnthropicModel`) based on `ModelProvider`
    - Build the agent via `LlmAgentBuilder::new("excel_agent")` with model, system instruction, and MCP toolset
    - Implement `system_instruction(artifact_dir)` function that includes: agent role as Excel specialist, available tool categories, workflow guidance (create/open → operate → save), artifact directory path, uploaded file path pattern `{artifact_dir}/{session_id}/{filename}`, write-only limitation note, and workbook TTL expiry note
    - _Requirements: 1.3, 2.1, 2.2, 2.3, 2.4, 3.1, 3.2, 3.3, 3.4, 3.5, 3.6, 3.7, 18.1_

- [ ] 3. Implement custom backend routes
  - [ ] 3.1 Implement AppError type and upload handler (`src/routes/upload.rs`)
    - Define `AppError` enum with `BadRequest`, `NotFound`, `PayloadTooLarge`, `Internal` variants, implementing `IntoResponse` for axum
    - Define `UploadResponse` and `ErrorResponse` structs
    - Implement `handle_upload` handler for `POST /api/upload/{session_id}`: accept multipart form data, validate file extension against allowed list (xlsx, xlsm, xls, ods), enforce 50 MB size limit, save to `{artifact_dir}/{session_id}/{filename}`, return path and confirmation message
    - Return HTTP 400 for bad extension (listing supported formats), HTTP 413 for oversized files
    - _Requirements: 6.1, 6.2, 6.3, 6.4, 6.5, 6.6_

  - [ ]* 3.2 Write property tests for upload validation
    - **Property 3: File extension validation**
    - **Validates: Requirements 6.2, 6.6**

  - [ ]* 3.3 Write property tests for upload storage
    - **Property 4: Upload stores file and returns correct path**
    - **Validates: Requirements 6.3, 6.4**

  - [ ]* 3.4 Write property test for upload size limit
    - **Property 5: Upload size limit enforcement**
    - **Validates: Requirements 6.5**

  - [ ] 3.5 Implement artifact serving (`src/routes/artifacts.rs`)
    - Implement `serve_artifact` handler for `GET /api/artifacts/{session_id}/{filename}`: read file from artifact directory, set correct content-type header based on extension (xlsx, xlsm, xls, ods), set Content-Disposition for download
    - Return HTTP 404 with descriptive message if artifact does not exist
    - Log a warning if the requested path is outside the expected artifact directory
    - _Requirements: 5.1, 5.2, 5.3, 5.4, 5.5, 18.3, 18.4_

  - [ ]* 3.6 Write property test for artifact round-trip
    - **Property 15: Artifact round trip — saved files are downloadable**
    - **Validates: Requirements 18.3**

  - [ ] 3.7 Implement static file serving (`src/static_files.rs`)
    - Implement `serve(static_dir)` that returns an axum Router serving files from the configured directory
    - Implement SPA fallback: serve `index.html` for all non-API routes that don't match a static file
    - Log a warning and continue in API-only mode if the static directory does not exist
    - _Requirements: 16.1, 16.2, 16.3_

  - [ ]* 3.8 Write property test for SPA fallback
    - **Property 14: Non-API routes serve index.html (SPA fallback)**
    - **Validates: Requirements 16.2**

- [ ] 4. Wire backend main entry point
  - [ ] 4.1 Implement main.rs server startup
    - Initialize `tracing_subscriber` with stderr output
    - Parse `AppConfig::from_env()`, exit with error on failure
    - Spawn MCP server via `spawn_mcp_server()`, exit with error on failure
    - Start MCP health monitoring background task
    - Build agent via `build_agent()`
    - Create `InMemorySessionService`
    - Configure `SecurityConfig` based on environment (development → `SecurityConfig::development()`, production → `SecurityConfig::production(origins)`)
    - Create `ServerConfig` with agent and session service, call `create_app()`
    - Merge custom routes: upload endpoint, artifact endpoint, static file fallback
    - Bind to configured host:port and serve
    - _Requirements: 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 4.1, 7.1, 7.2, 7.4, 8.1, 8.2, 17.1, 17.2, 17.3_

- [ ] 5. Checkpoint — Backend compiles and starts
  - Ensure the backend compiles, the MCP server spawns successfully, and the health endpoint responds. Ask the user if questions arise.

- [ ] 6. Initialize frontend project and implement core types
  - [ ] 6.1 Scaffold React project with Vite, TypeScript, and Tailwind CSS
    - Create `frontend/` directory, initialize with Vite React-TS template
    - Install dependencies: `react`, `react-dom`, `tailwindcss`, `@tailwindcss/vite`
    - Install dev dependencies: `vitest`, `@testing-library/react`, `@testing-library/jest-dom`, `fast-check`, `jsdom`
    - Configure Vite with API proxy to backend during development
    - Configure Tailwind CSS
    - Configure Vitest with jsdom environment
    - _Requirements: 9.4, 9.5_

  - [ ] 6.2 Define TypeScript types (`src/types/`)
    - Create `events.ts`: `SSEEvent`, `TextEvent`, `ToolCallEvent`, `ToolResultEvent`, `CompletionEvent`, `ErrorEvent` types
    - Create `spreadsheet.ts`: `SpreadsheetState`, `Sheet`, `Cell`, `CellFormat`, `ArtifactInfo` interfaces
    - Create `session.ts`: `SessionState` interface, localStorage key constants
    - _Requirements: 4.2, 4.3, 4.4, 11.1, 11.5_

  - [ ] 6.3 Implement utility modules (`src/utils/`)
    - Create `cellRef.ts`: `indexToColLetter()`, `colLetterToIndex()`, `parseCellRef()` functions for A1 notation conversion
    - Create `formatters.ts`: display helpers for tool arguments and cell values
    - _Requirements: 11.1_

  - [ ] 6.4 Implement API service (`src/services/api.ts`)
    - Implement `checkHealth()`: GET `/api/health`
    - Implement `createSession(userId)`: POST `/api/sessions`
    - Implement `listSessions()`: GET `/api/sessions`
    - Implement `uploadFile(sessionId, file)`: POST `/api/upload/{session_id}` with multipart form data and progress tracking via XMLHttpRequest
    - Implement `downloadArtifact(sessionId, filename)`: GET `/api/artifacts/{session_id}/{filename}`, trigger browser download
    - _Requirements: 7.2, 7.4, 8.1, 12.2, 13.2_

  - [ ] 6.5 Implement SSE parser (`src/services/sseParser.ts`)
    - Parse raw SSE stream lines into typed `SSEEvent` objects
    - Handle `text`, `tool-call`, `tool-result`, `completion`, and `error` event types
    - Extract `data` field from tool-result JSON strings, fall back to raw string on parse failure
    - _Requirements: 4.2, 4.3, 4.4, 4.5_

- [ ] 7. Implement frontend hooks
  - [ ] 7.1 Implement useSession hook (`src/hooks/useSession.ts`)
    - Generate UUID v4 via `crypto.randomUUID()` on first visit, persist to localStorage as `excel_agent_user_id`
    - Restore `userId` from localStorage on subsequent visits
    - Store and restore `currentSessionId` in localStorage as `excel_agent_session_id`
    - Expose `createNewSession()` that calls the API and updates state
    - On load, attempt to restore last active session ID
    - _Requirements: 7.5, 14.1, 14.2, 14.3, 14.4_

  - [ ]* 7.2 Write property test for session persistence
    - **Property 13: Session ID round-trips through localStorage**
    - **Validates: Requirements 14.3, 14.4**

  - [ ] 7.3 Implement useSSE hook (`src/hooks/useSSE.ts`)
    - Implement `sendMessage(message)` that POSTs to `/api/run/excel_agent/{userId}/{sessionId}` and reads the response body as a `ReadableStream`
    - Parse the SSE stream using `sseParser`, dispatch events to state
    - Track `isStreaming` state: true while stream is active, false on completion or error
    - Accumulate events array for consumption by chat and spreadsheet state
    - Handle fetch errors and stream interruptions: set `isStreaming = false`, append error event
    - _Requirements: 4.1, 10.3, 10.4, 10.6, 10.7, 10.8_

  - [ ]* 7.4 Write property test for text event concatenation
    - **Property 6: Text event concatenation**
    - **Validates: Requirements 10.4**

  - [ ] 7.5 Implement useSpreadsheetState hook (`src/hooks/useSpreadsheetState.ts`)
    - Implement reducer that processes SSE events into `SpreadsheetState`
    - Handle `create_workbook`: initialize state with "Sheet1"
    - Handle `write_cells`: set cell values from the `cells` array argument
    - Handle `write_row`: set consecutive cell values in a row from `start_cell`
    - Handle `write_column`: set consecutive cell values in a column from `start_cell`
    - Handle `set_cell_format`: apply format properties to cells in the range
    - Handle `merge_cells`: add range to `mergedRanges`
    - Handle `add_sheet`, `rename_sheet`, `delete_sheet`: mutate sheets map
    - Handle `set_column_width`, `set_row_height`: update layout maps
    - Handle `save_workbook` tool-result: extract filename, set `savedArtifact`
    - Handle `open_workbook` tool-result: parse sheet info, initialize sheets
    - _Requirements: 11.2, 11.3, 11.4_

  - [ ]* 7.6 Write property test for write tool state updates
    - **Property 9: Write tool events update preview state**
    - **Validates: Requirements 11.2, 11.3**

  - [ ] 7.7 Implement useFileUpload hook (`src/hooks/useFileUpload.ts`)
    - Implement `uploadFile(file)` that sends multipart form data to the upload endpoint
    - Track upload progress via XMLHttpRequest `progress` event
    - Return upload result (path, filename) or error
    - _Requirements: 13.2, 13.4, 13.5_

- [ ] 8. Checkpoint — Frontend hooks compile and unit tests pass
  - Ensure all hooks compile, types are consistent, and any written tests pass. Ask the user if questions arise.

- [ ] 9. Implement frontend components — App shell and layout
  - [ ] 9.1 Implement SplitPane component (`src/components/common/SplitPane.tsx`)
    - Render two children side by side with a draggable divider
    - Track divider position in state, apply as flex-basis percentages
    - _Requirements: 9.1, 9.2_

  - [ ] 9.2 Implement ErrorBanner component (`src/components/common/ErrorBanner.tsx`)
    - Display a connection error banner with a retry button
    - On retry click, call `checkHealth()` and dismiss on success
    - _Requirements: 15.1, 15.3_

  - [ ] 9.3 Implement Toast component (`src/components/common/Toast.tsx`)
    - Render toast notifications for transient errors
    - Auto-dismiss after 5 seconds
    - Support multiple simultaneous toasts
    - _Requirements: 15.4_

  - [ ] 9.4 Implement Toolbar component (`src/components/Toolbar.tsx`)
    - Display application title, "New Session" button, and model provider indicator
    - Wire "New Session" button to `useSession.createNewSession()`
    - _Requirements: 9.3, 14.1_

  - [ ] 9.5 Implement App shell (`src/App.tsx`)
    - Compose Toolbar + SplitPane(ChatPanel, PreviewPanel)
    - Initialize useSession, useSSE, useSpreadsheetState, useFileUpload hooks
    - On mount, check health endpoint; show ErrorBanner if unreachable
    - Pass hook state down to child components
    - _Requirements: 9.1, 9.2, 9.3, 15.1_

- [ ] 10. Implement Chat Panel components
  - [ ] 10.1 Implement ChatInput component (`src/components/ChatPanel/ChatInput.tsx`)
    - Render text input area with send button and file upload button (attachment icon)
    - Disable input and show loading indicator while `isStreaming` is true
    - On send: call `useSSE.sendMessage()` with the input text
    - On file select: call `useFileUpload.uploadFile()`
    - _Requirements: 10.2, 10.7, 13.1_

  - [ ] 10.2 Implement ToolCard component (`src/components/ChatPanel/ToolCard.tsx`)
    - Render a collapsible card showing tool name, arguments (formatted), and result when available
    - Default to collapsed; expand on click
    - For `save_workbook` tool-result with success, render a download button
    - _Requirements: 10.5, 12.1_

  - [ ]* 10.3 Write property test for tool card rendering
    - **Property 7: Tool-call events produce complete tool cards**
    - **Validates: Requirements 10.5**

  - [ ]* 10.4 Write property test for download button on save
    - **Property 12: Successful save produces download button**
    - **Validates: Requirements 12.1**

  - [ ] 10.5 Implement message components (`UserMessage.tsx`, `AgentMessage.tsx`, `MessageList.tsx`)
    - `UserMessage`: render user message bubble
    - `AgentMessage`: render streaming text content + inline ToolCards + download button when applicable
    - `MessageList`: scrollable container, auto-scroll to bottom on new messages
    - _Requirements: 10.1, 10.4, 10.5, 10.6_

  - [ ] 10.6 Implement ChatPanel container (`src/components/ChatPanel/ChatPanel.tsx`)
    - Compose MessageList + ChatInput
    - On upload success: insert a system message with file path and prompt to instruct the agent
    - On SSE error: display error message in chat, re-enable input
    - On SSE stream disconnect mid-response: display warning message
    - _Requirements: 10.1, 10.2, 10.3, 10.8, 13.3, 15.2_

- [ ] 11. Implement Preview Panel components
  - [ ] 11.1 Implement CellRenderer component (`src/components/PreviewPanel/CellRenderer.tsx`)
    - Render a table cell with the cell's value
    - Apply CSS styles from `CellFormat`: `font-weight: bold`, `font-style: italic`, `color` from `fontColor`, `background-color` from `backgroundColor`
    - _Requirements: 11.5_

  - [ ]* 11.2 Write property test for cell formatting
    - **Property 11: Cell formatting reflected in rendering**
    - **Validates: Requirements 11.5**

  - [ ] 11.3 Implement SpreadsheetGrid component (`src/components/PreviewPanel/SpreadsheetGrid.tsx`)
    - Render column headers (A, B, C...) and row numbers
    - Compute grid bounds from populated cells
    - Render cells using CellRenderer, with sticky headers
    - Support virtual rendering for large grids (render only visible rows/columns based on scroll position)
    - _Requirements: 11.1_

  - [ ]* 11.4 Write property test for grid rendering
    - **Property 8: Spreadsheet grid renders all populated cells**
    - **Validates: Requirements 11.1**

  - [ ] 11.5 Implement SheetTabs component (`src/components/PreviewPanel/SheetTabs.tsx`)
    - Render tab bar with one tab per sheet
    - Highlight active sheet tab
    - On tab click, switch `activeSheet` in state
    - Only render when sheet count > 1
    - _Requirements: 11.4_

  - [ ]* 11.6 Write property test for sheet tab count
    - **Property 10: Sheet tabs match sheet count**
    - **Validates: Requirements 11.4**

  - [ ] 11.7 Implement PreviewPanel container (`src/components/PreviewPanel/PreviewPanel.tsx`)
    - Compose toolbar (with download button when `savedArtifact` exists + limitations notice) + SpreadsheetGrid + SheetTabs
    - Show empty state message when no workbook has been created ("Ask the agent to create a spreadsheet")
    - Display persistent notice about features not rendered in preview (charts, images, tables, conditional formatting, data validation, sparklines)
    - Wire download button to `api.downloadArtifact()`
    - _Requirements: 11.6, 11.7, 12.2, 12.3, 12.4_

- [ ] 12. Checkpoint — Frontend renders and all component tests pass
  - Ensure the frontend builds, renders the split-pane layout, and all written tests pass. Ask the user if questions arise.

- [ ] 13. Integration wiring and final assembly
  - [ ] 13.1 Wire frontend to backend SSE stream end-to-end
    - Ensure `useSSE` correctly connects to the backend SSE endpoint
    - Verify text events stream into AgentMessage, tool-call/tool-result events render as ToolCards, and completion event re-enables input
    - Verify spreadsheet state updates from tool events flow into PreviewPanel
    - _Requirements: 4.1, 4.2, 4.3, 4.4, 4.5, 10.3, 10.4, 10.5, 10.6, 11.2, 11.3_

  - [ ] 13.2 Wire file upload and download flows
    - Verify upload button → multipart POST → system message in chat flow
    - Verify download button → artifact GET → browser download flow
    - Verify PreviewPanel toolbar download button works when `savedArtifact` is set
    - _Requirements: 6.1, 6.3, 6.4, 12.1, 12.2, 12.3, 13.1, 13.2, 13.3_

  - [ ] 13.3 Wire session management
    - Verify "New Session" creates session via API, clears chat and preview, updates localStorage
    - Verify page refresh restores session ID and user ID from localStorage
    - Verify session ID is used in all SSE and upload requests
    - _Requirements: 7.1, 7.2, 7.3, 7.5, 14.1, 14.2, 14.3, 14.4_

  - [ ] 13.4 Configure frontend production build and static serving
    - Ensure `vite build` outputs to `frontend/dist`
    - Verify the backend serves the built frontend at non-API routes
    - Verify SPA fallback works for client-side routes
    - _Requirements: 16.1, 16.2, 16.3_

- [ ] 14. Final checkpoint — Full application works end-to-end
  - Ensure all tests pass, the backend starts and serves the frontend, SSE streaming works, file upload/download works, and session management persists across refreshes. Ask the user if questions arise.

## Notes

- Tasks marked with `*` are optional and can be skipped for faster MVP
- Each task references specific requirements for traceability
- Checkpoints ensure incremental validation
- Property tests validate universal correctness properties from the design document
- The excel-mcp-server binary must be built first — configure its path via `MCP_SERVER_PATH` environment variable
- All adk-rust crates are referenced via path dependencies pointing to `~/Developer/projects/adk-rust`
