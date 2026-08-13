use std::sync::Arc;

use excel_mcp_server::{server::ExcelMcpServer, store::WorkbookStore};
use rmcp::{
    ClientLifecycleMode, ClientServiceExt, ServiceExt,
    model::{CacheScope, CallToolRequestParams, ProtocolVersion},
};
use serde_json::json;
use tokio::sync::RwLock;

async fn connect(
    store: Arc<RwLock<WorkbookStore>>,
) -> (
    rmcp::service::RunningService<rmcp::RoleClient, ()>,
    rmcp::service::RunningService<rmcp::RoleServer, ExcelMcpServer>,
) {
    let (server_transport, client_transport) = tokio::io::duplex(64 * 1024);
    let server_handler = ExcelMcpServer::new(store);
    let (server, client) = tokio::join!(
        server_handler.serve(server_transport),
        ().serve_with_lifecycle(
            client_transport,
            ClientLifecycleMode::Discover {
                preferred_versions: vec![ProtocolVersion::V_2026_07_28],
            },
        ),
    );
    (
        client.expect("discover server"),
        server.expect("start server"),
    )
}

#[tokio::test]
async fn workbook_handle_survives_a_new_stateless_handler() {
    let store = Arc::new(RwLock::new(WorkbookStore::new()));

    let (first_client, first_server) = connect(Arc::clone(&store)).await;
    let created = first_client
        .call_tool(CallToolRequestParams::new("create_workbook").with_arguments(Default::default()))
        .await
        .expect("create workbook");
    let created_text = created.content[0]
        .as_text()
        .expect("text result")
        .text
        .clone();
    let created_json: serde_json::Value = serde_json::from_str(&created_text).expect("JSON result");
    let workbook_id = created_json["data"]["workbook_id"]
        .as_str()
        .expect("workbook id")
        .to_string();
    first_client.cancellation_token().cancel();
    first_server.cancellation_token().cancel();

    // A new handler represents a later stateless HTTP request which carries no
    // session id and may be processed by a different service object.
    let (second_client, second_server) = connect(Arc::clone(&store)).await;
    let tools = second_client
        .list_tools(None)
        .await
        .expect("list tools with cache hint");
    assert_eq!(tools.ttl_ms, Some(3_600_000));
    assert_eq!(tools.cache_scope, Some(CacheScope::Public));
    let listed = second_client
        .call_tool(
            CallToolRequestParams::new("list_sheets").with_arguments(
                json!({"workbook_id": workbook_id})
                    .as_object()
                    .unwrap()
                    .clone(),
            ),
        )
        .await
        .expect("list sheets through a new handler");
    let listed_text = listed.content[0]
        .as_text()
        .expect("text result")
        .text
        .clone();
    let listed_json: serde_json::Value = serde_json::from_str(&listed_text).expect("JSON result");
    assert_eq!(listed_json["status"], "success");
    assert_eq!(listed_json["data"][0]["name"], "Sheet1");

    second_client.cancellation_token().cancel();
    second_server.cancellation_token().cancel();
}
