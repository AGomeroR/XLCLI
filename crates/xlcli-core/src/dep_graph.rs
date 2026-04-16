use std::collections::VecDeque;

use petgraph::graphmap::DiGraphMap;
use petgraph::Direction;

use crate::types::CellAddr;

pub struct DepGraph {
    graph: DiGraphMap<CellAddr, ()>,
}

impl DepGraph {
    pub fn new() -> Self {
        Self {
            graph: DiGraphMap::new(),
        }
    }

    pub fn set_dependencies(&mut self, cell: CellAddr, deps: Vec<CellAddr>) {
        let old_deps: Vec<CellAddr> = self
            .graph
            .neighbors_directed(cell, Direction::Outgoing)
            .collect();
        for dep in old_deps {
            self.graph.remove_edge(cell, dep);
        }

        if !self.graph.contains_node(cell) && !deps.is_empty() {
            self.graph.add_node(cell);
        }

        for dep in deps {
            if !self.graph.contains_node(dep) {
                self.graph.add_node(dep);
            }
            self.graph.add_edge(cell, dep, ());
        }
    }

    pub fn remove_cell(&mut self, cell: CellAddr) {
        self.graph.remove_node(cell);
    }

    pub fn dependents_toposorted(&self, changed: CellAddr) -> Vec<CellAddr> {
        let mut dirty = Vec::new();
        let mut queue = VecDeque::new();
        let mut visited = std::collections::HashSet::new();

        queue.push_back(changed);
        visited.insert(changed);

        while let Some(node) = queue.pop_front() {
            for dependent in self.graph.neighbors_directed(node, Direction::Incoming) {
                if visited.insert(dependent) {
                    dirty.push(dependent);
                    queue.push_back(dependent);
                }
            }
        }

        self.toposort_subset(&dirty)
    }

    fn toposort_subset(&self, cells: &[CellAddr]) -> Vec<CellAddr> {
        let cell_set: std::collections::HashSet<CellAddr> = cells.iter().copied().collect();
        let mut in_degree: std::collections::HashMap<CellAddr, usize> = std::collections::HashMap::new();

        for &cell in cells {
            let deps_in_set = self
                .graph
                .neighbors_directed(cell, Direction::Outgoing)
                .filter(|dep| cell_set.contains(dep))
                .count();
            in_degree.insert(cell, deps_in_set);
        }

        let mut queue: VecDeque<CellAddr> = in_degree
            .iter()
            .filter(|(_, &deg)| deg == 0)
            .map(|(&cell, _)| cell)
            .collect();

        let mut sorted = Vec::with_capacity(cells.len());

        while let Some(cell) = queue.pop_front() {
            sorted.push(cell);
            for dependent in self.graph.neighbors_directed(cell, Direction::Incoming) {
                if let Some(deg) = in_degree.get_mut(&dependent) {
                    *deg = deg.saturating_sub(1);
                    if *deg == 0 {
                        queue.push_back(dependent);
                    }
                }
            }
        }

        sorted
    }

    pub fn has_cycle(&self, cell: CellAddr) -> bool {
        let mut visited = std::collections::HashSet::new();
        let mut stack = vec![cell];

        while let Some(node) = stack.pop() {
            if !visited.insert(node) {
                if node == cell {
                    return true;
                }
                continue;
            }
            for dep in self.graph.neighbors_directed(node, Direction::Outgoing) {
                if dep == cell {
                    return true;
                }
                stack.push(dep);
            }
        }
        false
    }
}

impl Default for DepGraph {
    fn default() -> Self {
        Self::new()
    }
}
