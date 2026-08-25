use std::time::Duration;

use rmcp::model::ProtocolVersion;
use rmcp::transport::TokioChildProcess;
use rmcp::{ClientLifecycleMode, ClientServiceExt};
use tokio::process::Command;

fn server_transport() -> TokioChildProcess {
    TokioChildProcess::new(Command::new(env!("CARGO_BIN_EXE_excel-mcp-server")))
        .expect("spawn excel MCP server")
}

#[tokio::test]
async fn supports_2026_server_discover_lifecycle() {
    let client = tokio::time::timeout(
        Duration::from_secs(10),
        ().serve_with_lifecycle(
            server_transport(),
            ClientLifecycleMode::Auto {
                preferred_versions: vec![ProtocolVersion::V_2026_07_28],
                legacy_version: Some(ProtocolVersion::V_2025_11_25),
            },
        ),
    )
    .await
    .expect("2026 handshake timed out")
    .expect("2026 handshake failed");

    let tools = client
        .list_tools(Default::default())
        .await
        .expect("list tools after discover");
    assert!(
        tools.tools.len() >= 90,
        "unexpected tool count: {}",
        tools.tools.len()
    );
    client.cancel().await.expect("stop server");
}

#[tokio::test]
async fn retains_legacy_initialize_lifecycle() {
    let client = tokio::time::timeout(
        Duration::from_secs(10),
        ().serve_with_lifecycle(server_transport(), ClientLifecycleMode::Initialize),
    )
    .await
    .expect("legacy handshake timed out")
    .expect("legacy handshake failed");

    let tools = client
        .list_tools(Default::default())
        .await
        .expect("list tools after initialize");
    assert!(
        tools.tools.len() >= 90,
        "unexpected tool count: {}",
        tools.tools.len()
    );
    client.cancel().await.expect("stop server");
}
