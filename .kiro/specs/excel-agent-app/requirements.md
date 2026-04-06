# Requirements Document

## Introduction

This document specifies the requirements for a full-stack AI agent application that enables users to create, edit, and style Excel documents through natural language conversation. The application consists of a Rust backend powered by the adk-rust framework (v0.5.0) that orchestrates an LLM agent connected to the Excel MCP server, and a React frontend that provides a chat interface with real-time spreadsheet preview and file management capabilities. The backend spawns the Excel MCP server as a child process, exposes the agent over HTTP with SSE streaming, and manages session-scoped artifacts (generated Excel files). The frontend consumes the SSE stream to display agent responses and renders a live preview of the spreadsheet as it is constructed.

## Glossary

- **Agent_Server**: The Rust backend process built with adk-rust that hosts the LLM agent, manages sessions, and serves the HTTP API and static frontend assets
- **Excel_Agent**: The LLM-powered agent created via `LlmAgentBuilder` that receives user messages, reasons about them, and invokes Excel MCP tools to fulfill requests
- **Excel_MCP_Server**: The sibling MCP server binary (from `.kiro/specs/excel-mcp-server/`) that exposes 25+ Excel manipulation tools over stdio transport
- **MCP_Toolset**: The `McpToolset` instance that connects the Excel_Agent to the Excel_MCP_Server's tools via the spawned child process
- **Session**: An adk-rust session identified by a user ID and session ID, maintaining conversation history and associated artifacts
- **Artifact**: A file (typically a generated .xlsx) stored by the adk-server artifact service, scoped to a specific Session
- **SSE_Stream**: The Server-Sent Events stream provided by adk-server at `/api/run/{app_name}/{user_id}/{session_id}` for real-time agent response delivery
- **Chat_Panel**: The React UI component where users type messages and view the conversation history with the Excel_Agent
- **Preview_Panel**: The React UI component that renders a visual spreadsheet representation of the Excel file being constructed or edited
- **Model_Provider**: A configured LLM backend (Gemini, OpenAI, or Anthropic) that the Excel_Agent uses for reasoning, selectable via environment configuration

## Requirements

### Requirement 1: Backend Server Initialization

**User Story:** As a developer, I want the Agent_Server to start up and initialize all dependencies, so that the application is ready to accept user requests.

#### Acceptance Criteria

1. WHEN the Agent_Server process starts, THE Agent_Server SHALL spawn the Excel_MCP_Server as a child process using `TokioChildProcess` and establish an MCP connection over stdio
2. WHEN the MCP connection is established, THE Agent_Server SHALL create an MCP_Toolset from the connected MCP client and discover all available Excel tools
3. WHEN the MCP_Toolset is ready, THE Agent_Server SHALL build an Excel_Agent using `LlmAgentBuilder` with the configured Model_Provider, a system instruction describing the agent's Excel specialist role, and the MCP_Toolset
4. WHEN the Excel_Agent is built, THE Agent_Server SHALL create a `ServerConfig` with the agent, an `InMemorySessionService`, and a `SecurityConfig`, and bind the HTTP server to the configured host and port
5. IF the Excel_MCP_Server child process fails to start, THEN THE Agent_Server SHALL log the error and exit with a non-zero exit code
6. IF the Model_Provider configuration is missing or invalid, THEN THE Agent_Server SHALL log a descriptive error identifying the expected environment variables and exit with a non-zero exit code

### Requirement 2: Model Provider Configuration

**User Story:** As a developer, I want to choose which LLM provider powers the Excel_Agent, so that I can use whichever model best fits my needs and budget.

#### Acceptance Criteria

1. THE Agent_Server SHALL read the model provider selection from the `MODEL_PROVIDER` environment variable, accepting values "gemini", "openai", or "anthropic"
2. WHEN `MODEL_PROVIDER` is set to "gemini", THE Agent_Server SHALL configure a `GeminiModel` using the `GOOGLE_API_KEY` environment variable and the model name from `GEMINI_MODEL` (default: "gemini-2.5-flash")
3. WHEN `MODEL_PROVIDER` is set to "openai", THE Agent_Server SHALL configure an `OpenAiModel` using the `OPENAI_API_KEY` environment variable and the model name from `OPENAI_MODEL` (default: "gpt-4o")
4. WHEN `MODEL_PROVIDER` is set to "anthropic", THE Agent_Server SHALL configure an `AnthropicModel` using the `ANTHROPIC_API_KEY` environment variable and the model name from `ANTHROPIC_MODEL` (default: "claude-sonnet-4-20250514")
5. IF the required API key environment variable for the selected provider is not set, THEN THE Agent_Server SHALL log an error naming the missing variable and exit with a non-zero exit code

### Requirement 3: Agent System Instruction

**User Story:** As a user, I want the Excel_Agent to understand its role and capabilities, so that it provides accurate and helpful responses about Excel file creation and editing.

#### Acceptance Criteria

1. THE Agent_Server SHALL configure the Excel_Agent with a system instruction that describes the agent as an Excel specialist capable of creating, editing, formatting, and analyzing Excel files
2. THE Agent_Server SHALL include in the system instruction a summary of available tool categories: workbook lifecycle, reading data, writing data, formatting, charts, images, tables, conditional formatting, data validation, layout controls, sparklines, and search
3. THE Agent_Server SHALL include in the system instruction guidance on the workflow: create or open a workbook first, then perform operations, then save the workbook to produce a downloadable file
4. THE Agent_Server SHALL include in the system instruction the constraint that the agent must always save completed workbooks to the artifact directory so they are available for download
5. THE Agent_Server SHALL include in the system instruction a note that workbook handles may expire after 30 minutes of inactivity due to the MCP server's TTL eviction, and that the agent should reopen the file if a workbook_id becomes invalid
6. THE Agent_Server SHALL include in the system instruction the path pattern for uploaded files (`{artifact_dir}/{session_id}/{filename}`) so the agent knows where to find files the user has uploaded
7. THE Agent_Server SHALL include in the system instruction a note that newly created workbooks (via `create_workbook`) cannot be read back — the agent must track its own writes and should not attempt to call `read_sheet` or `read_cell` on them

### Requirement 4: SSE Streaming API

**User Story:** As a frontend developer, I want to consume real-time agent responses via SSE, so that the Chat_Panel can display the agent's work as it happens.

#### Acceptance Criteria

1. THE Agent_Server SHALL expose the adk-server SSE endpoint at `/api/run/{app_name}/{user_id}/{session_id}` that accepts POST requests with a user message and returns an SSE stream
2. WHEN the Excel_Agent produces text content during execution, THE Agent_Server SHALL stream each content chunk as an SSE event to the connected client
3. WHEN the Excel_Agent invokes an MCP tool, THE Agent_Server SHALL stream a tool-call event containing the tool name and arguments to the connected client
4. WHEN the Excel_Agent receives a tool result, THE Agent_Server SHALL stream a tool-result event containing the tool response to the connected client
5. WHEN the agent execution completes, THE Agent_Server SHALL send a completion SSE event and close the stream

### Requirement 5: Artifact Management for Generated Files

**User Story:** As a user, I want the Excel files I create through the agent to be stored and downloadable, so that I can retrieve the finished spreadsheets.

#### Acceptance Criteria

1. THE Agent_Server SHALL configure the adk-server with an artifact service that stores files scoped to each Session
2. THE Agent_Server SHALL configure a dedicated artifact directory where the Excel_Agent saves generated workbook files
3. WHEN the Excel_Agent saves a workbook via the `save_workbook` MCP tool, THE Agent_Server SHALL ensure the file is saved to the session's artifact directory
4. THE Agent_Server SHALL expose an artifact retrieval endpoint that allows the frontend to download stored Excel files by artifact identifier
5. IF an artifact is requested that does not exist, THEN THE Agent_Server SHALL return an HTTP 404 response with a descriptive error message

### Requirement 6: File Upload for Editing Existing Workbooks

**User Story:** As a user, I want to upload an existing Excel file, so that the Excel_Agent can read, analyze, or edit it on my behalf.

#### Acceptance Criteria

1. THE Agent_Server SHALL expose an HTTP POST endpoint at `/api/upload/{session_id}` for uploading Excel files, accepting multipart form data with a file field
2. WHEN a file is uploaded, THE Agent_Server SHALL validate that the file extension is one of: .xlsx, .xlsm, .xls, .ods
3. WHEN a valid file is uploaded, THE Agent_Server SHALL store the file in the session's artifact directory and return the stored file path
4. WHEN a valid file is uploaded, THE Agent_Server SHALL return a confirmation message that includes the file path, enabling the user to instruct the Excel_Agent to open it
5. IF the uploaded file exceeds 50 MB, THEN THE Agent_Server SHALL reject the upload with an HTTP 413 response and a message stating the maximum allowed file size
6. IF the uploaded file has an unsupported extension, THEN THE Agent_Server SHALL reject the upload with an HTTP 400 response listing the supported formats

### Requirement 7: Session Management

**User Story:** As a user, I want my conversation and files to persist within a session, so that I can have multi-turn interactions with the Excel_Agent without losing context.

#### Acceptance Criteria

1. THE Agent_Server SHALL use `InMemorySessionService` to maintain conversation history per Session
2. WHEN a new session is created via the `/api/sessions` endpoint, THE Agent_Server SHALL return a session ID that the frontend uses for all subsequent requests
3. WHEN a user sends a message within an existing Session, THE Agent_Server SHALL include the full conversation history as context for the Excel_Agent
4. THE Agent_Server SHALL expose the adk-server session listing endpoint at `/api/sessions` for the frontend to enumerate active sessions
5. THE Frontend SHALL generate a unique user ID (UUID v4) on first visit and persist it in browser localStorage, reusing it for all subsequent sessions

### Requirement 8: Health and Discovery Endpoints

**User Story:** As a frontend developer, I want health and app discovery endpoints, so that the UI can verify the backend is running and discover available agents.

#### Acceptance Criteria

1. THE Agent_Server SHALL expose the adk-server health endpoint at `/api/health` that returns HTTP 200 when the server is operational
2. THE Agent_Server SHALL expose the adk-server app listing endpoint at `/api/apps` that returns metadata about the Excel_Agent including its name and description
3. WHEN the Excel_MCP_Server child process has crashed, THE Agent_Server SHALL report a degraded status through the health endpoint
4. WHEN the Excel_MCP_Server child process exits unexpectedly, THE Agent_Server SHALL attempt to respawn the child process and re-establish the MCP connection, logging the restart attempt to stderr
5. IF the Excel_MCP_Server child process fails to restart after 3 consecutive attempts, THEN THE Agent_Server SHALL report a critical status through the health endpoint and return error responses for all tool-invoking requests until the process is manually restarted


### Requirement 9: React Frontend Application Shell

**User Story:** As a user, I want a web application with a clean layout, so that I can interact with the Excel_Agent and see my spreadsheet in a single view.

#### Acceptance Criteria

1. THE Frontend SHALL render a split-pane layout with the Chat_Panel on the left and the Preview_Panel on the right
2. THE Frontend SHALL allow the user to resize the split between the Chat_Panel and Preview_Panel by dragging a divider
3. THE Frontend SHALL display a toolbar at the top containing the application title, session controls, and a model provider indicator
4. THE Frontend SHALL be built with React 18+, TypeScript, and Vite as the build tool
5. THE Frontend SHALL use Tailwind CSS for styling

### Requirement 10: Chat Panel

**User Story:** As a user, I want to type messages and see the agent's responses in a conversational interface, so that I can describe what Excel file I want and follow the agent's progress.

#### Acceptance Criteria

1. THE Chat_Panel SHALL display a scrollable message list showing user messages and agent responses in chronological order
2. THE Chat_Panel SHALL provide a text input area at the bottom with a send button for submitting messages to the Excel_Agent
3. WHEN the user submits a message, THE Chat_Panel SHALL send a POST request to the SSE_Stream endpoint and begin consuming the streamed response
4. WHILE the SSE_Stream is active, THE Chat_Panel SHALL append streamed text content to the current agent message in real time
5. WHILE the SSE_Stream is active, THE Chat_Panel SHALL display tool invocations as collapsible cards showing the tool name, arguments, and result
6. WHEN the SSE_Stream completes, THE Chat_Panel SHALL mark the agent message as complete and re-enable the input area
7. WHILE the SSE_Stream is active, THE Chat_Panel SHALL disable the input area and display a loading indicator
8. IF the SSE_Stream connection fails, THEN THE Chat_Panel SHALL display an error message and re-enable the input area

### Requirement 11: Spreadsheet Preview Panel

**User Story:** As a user, I want to see a visual preview of the Excel file as it is being built, so that I can verify the agent is creating what I asked for.

#### Acceptance Criteria

1. THE Preview_Panel SHALL render a grid-based spreadsheet view showing cell values, column headers (A, B, C...), and row numbers
2. WHEN the Excel_Agent completes a tool call that modifies the workbook (write_cells, write_row, write_column, set_cell_format, merge_cells, add_sheet), THE Preview_Panel SHALL update the displayed grid to reflect the changes
3. THE Preview_Panel SHALL parse tool-call and tool-result SSE events to extract spreadsheet state changes
4. THE Preview_Panel SHALL display multiple sheet tabs when the workbook contains more than one sheet, allowing the user to switch between sheets
5. THE Preview_Panel SHALL render basic cell formatting including bold text, italic text, font color, and background color when formatting information is available from tool results
6. WHEN no workbook has been created yet in the current Session, THE Preview_Panel SHALL display an empty state message prompting the user to ask the agent to create a spreadsheet
7. THE Preview_Panel SHALL display a notice that charts, images, tables, conditional formatting, data validation, and sparklines are not rendered in the preview but will be present in the downloaded Excel file

### Requirement 12: File Download

**User Story:** As a user, I want to download the generated Excel file, so that I can use it in Excel, Google Sheets, or other spreadsheet applications.

#### Acceptance Criteria

1. WHEN the Excel_Agent saves a workbook and the tool result indicates success, THE Chat_Panel SHALL display a download button inline with the agent's response
2. WHEN the user clicks the download button, THE Frontend SHALL fetch the artifact from the Agent_Server's artifact endpoint and trigger a browser file download with the original filename
3. THE Preview_Panel SHALL display a persistent download button in its toolbar when a saved workbook artifact exists for the current Session
4. IF the artifact download request fails, THEN THE Frontend SHALL display an error notification with the failure reason

### Requirement 13: File Upload UI

**User Story:** As a user, I want to upload an existing Excel file through the UI, so that I can ask the agent to analyze or edit it.

#### Acceptance Criteria

1. THE Chat_Panel SHALL provide a file upload button (attachment icon) next to the message input area
2. WHEN the user selects a file via the upload button, THE Frontend SHALL send the file to the Agent_Server's upload endpoint as multipart form data
3. WHEN the upload succeeds, THE Frontend SHALL insert a system message in the Chat_Panel indicating the file was uploaded and its path, and prompt the user to instruct the agent to open it
4. WHILE a file upload is in progress, THE Frontend SHALL display an upload progress indicator
5. IF the upload fails, THEN THE Frontend SHALL display an error message with the reason from the server response

### Requirement 14: Session Management UI

**User Story:** As a user, I want to create new sessions and switch between them, so that I can work on multiple Excel projects independently.

#### Acceptance Criteria

1. THE Frontend SHALL display a "New Session" button in the toolbar that creates a new Session via the `/api/sessions` endpoint
2. WHEN a new Session is created, THE Frontend SHALL clear the Chat_Panel and Preview_Panel and begin using the new session ID for all subsequent requests
3. THE Frontend SHALL persist the current session ID in browser local storage so that refreshing the page resumes the same Session
4. WHEN the application loads, THE Frontend SHALL attempt to restore the last active session ID from local storage

### Requirement 15: Error Handling and Connection Resilience

**User Story:** As a user, I want clear feedback when something goes wrong, so that I understand what happened and can take corrective action.

#### Acceptance Criteria

1. IF the Frontend cannot reach the Agent_Server on initial load, THEN THE Frontend SHALL display a connection error banner with a retry button
2. IF the SSE_Stream disconnects mid-response, THEN THE Frontend SHALL display a warning message in the Chat_Panel indicating the response was interrupted
3. WHEN the user clicks the retry button on a connection error, THE Frontend SHALL attempt to reconnect to the Agent_Server health endpoint
4. THE Frontend SHALL display toast notifications for transient errors (upload failures, download failures, session creation failures) that auto-dismiss after 5 seconds

### Requirement 16: Static Asset Serving

**User Story:** As a developer, I want the Agent_Server to serve the React frontend as static files, so that the entire application runs from a single process in production.

#### Acceptance Criteria

1. THE Agent_Server SHALL serve the built React frontend static assets from a configurable directory (default: `./frontend/dist`)
2. THE Agent_Server SHALL serve `index.html` for all non-API routes to support client-side routing
3. WHEN the static asset directory does not exist, THE Agent_Server SHALL log a warning and continue running with only the API endpoints available

### Requirement 17: CORS and Security Configuration

**User Story:** As a developer, I want proper CORS and security configuration, so that the frontend can communicate with the backend during development and the application is secure in production.

#### Acceptance Criteria

1. WHEN the `ENVIRONMENT` variable is set to "development", THE Agent_Server SHALL use `SecurityConfig::development()` which allows all origins
2. WHEN the `ENVIRONMENT` variable is set to "production", THE Agent_Server SHALL use `SecurityConfig::production(origins)` with origins read from the `ALLOWED_ORIGINS` environment variable (comma-separated)
3. WHEN the `ENVIRONMENT` variable is not set, THE Agent_Server SHALL default to development mode and log a warning

### Requirement 18: Agent Artifact Save Workflow

**User Story:** As a user, I want the agent to automatically save files to a location the server can serve, so that I can download them without manual path management.

#### Acceptance Criteria

1. THE Agent_Server SHALL set an environment variable or include in the system instruction the artifact directory path where the Excel_Agent must save all workbooks
2. WHEN the Excel_Agent decides to save a workbook, THE Excel_Agent SHALL use the artifact directory path as the base directory for the `save_workbook` tool's file path argument
3. THE Agent_Server SHALL map saved workbook files in the artifact directory to downloadable artifact identifiers accessible via the artifact retrieval endpoint
4. IF the Excel_Agent attempts to save a workbook to a path outside the artifact directory, THEN THE Agent_Server SHALL still serve the file if it exists, but log a warning about the non-standard save location
