use std::collections::HashMap;
use std::time::{Duration, Instant};

use uuid::Uuid;

use crate::error::ExcelMcpError;

/// A single workbook in the store — always backed by zavora-xlsx.
pub struct WorkbookEntry {
    pub id: String,
    pub data: zavora_xlsx::Workbook,
    pub read_only: bool,
    pub last_access: Instant,
}

/// Manages the lifecycle of all open workbooks.
pub struct WorkbookStore {
    workbooks: HashMap<String, WorkbookEntry>,
    max_capacity: usize,
    ttl: Duration,
}

impl std::fmt::Debug for WorkbookStore {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("WorkbookStore")
            .field("open_count", &self.workbooks.len())
            .field("max_capacity", &self.max_capacity)
            .field("ttl", &self.ttl)
            .finish()
    }
}

impl WorkbookStore {
    pub fn new() -> Self {
        Self { workbooks: HashMap::new(), max_capacity: 10, ttl: Duration::from_secs(30 * 60) }
    }

    pub fn with_config(max_capacity: usize, ttl: Duration) -> Self {
        Self { workbooks: HashMap::new(), max_capacity, ttl }
    }

    pub fn insert(&mut self, mut entry: WorkbookEntry) -> Result<String, ExcelMcpError> {
        self.evict_expired();
        if self.workbooks.len() >= self.max_capacity {
            return Err(ExcelMcpError::CapacityExceeded(format!(
                "Workbook store is at maximum capacity ({}). Save and close an existing workbook first.",
                self.max_capacity
            )));
        }
        let id = Uuid::new_v4().to_string();
        entry.id = id.clone();
        entry.last_access = Instant::now();
        self.workbooks.insert(id.clone(), entry);
        Ok(id)
    }

    pub fn get(&mut self, id: &str) -> Option<&WorkbookEntry> {
        self.evict_expired();
        if let Some(entry) = self.workbooks.get_mut(id) { entry.last_access = Instant::now(); }
        self.workbooks.get(id)
    }

    pub fn get_mut(&mut self, id: &str) -> Option<&mut WorkbookEntry> {
        self.evict_expired();
        if let Some(entry) = self.workbooks.get_mut(id) { entry.last_access = Instant::now(); }
        self.workbooks.get_mut(id)
    }

    pub fn remove(&mut self, id: &str) -> Option<WorkbookEntry> { self.workbooks.remove(id) }

    pub fn evict_expired(&mut self) -> Vec<String> {
        let now = Instant::now();
        let ttl = self.ttl;
        let expired: Vec<String> = self.workbooks.iter()
            .filter(|(_, e)| now.duration_since(e.last_access) > ttl)
            .map(|(id, _)| id.clone()).collect();
        for id in &expired { self.workbooks.remove(id); }
        expired
    }

    pub fn open_ids(&self) -> Vec<String> { self.workbooks.keys().cloned().collect() }
    pub fn is_full(&self) -> bool { self.workbooks.len() >= self.max_capacity }
}

impl Default for WorkbookStore {
    fn default() -> Self { Self::new() }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_entry() -> WorkbookEntry {
        WorkbookEntry {
            id: String::new(),
            data: zavora_xlsx::Workbook::new(),
            read_only: false,
            last_access: Instant::now(),
        }
    }

    #[test]
    fn test_capacity_enforcement() {
        let mut store = WorkbookStore::with_config(2, Duration::from_secs(600));
        let _id1 = store.insert(make_entry()).unwrap();
        let _id2 = store.insert(make_entry()).unwrap();
        assert!(store.insert(make_entry()).is_err());
    }

    #[test]
    fn test_ttl_eviction() {
        let mut store = WorkbookStore::with_config(10, Duration::from_millis(1));
        let id = store.insert(make_entry()).unwrap();
        std::thread::sleep(Duration::from_millis(10));
        assert!(store.get(&id).is_none());
    }

    #[test]
    fn test_remove() {
        let mut store = WorkbookStore::with_config(10, Duration::from_secs(600));
        let id = store.insert(make_entry()).unwrap();
        assert!(store.remove(&id).is_some());
        assert!(store.get(&id).is_none());
    }
}
