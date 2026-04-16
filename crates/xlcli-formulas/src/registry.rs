use std::collections::HashMap;

use xlcli_core::cell::CellValue;

use crate::ast::Expr;
use crate::eval::EvalContext;

pub struct FnSpec {
    pub name: &'static str,
    pub min_args: usize,
    pub max_args: Option<usize>,
    pub eval: fn(&[Expr], &dyn EvalContext, &FunctionRegistry) -> CellValue,
}

pub struct FunctionRegistry {
    funcs: HashMap<&'static str, FnSpec>,
}

impl FunctionRegistry {
    pub fn new() -> Self {
        Self {
            funcs: HashMap::new(),
        }
    }

    pub fn register(&mut self, spec: FnSpec) {
        self.funcs.insert(spec.name, spec);
    }

    pub fn get(&self, name: &str) -> Option<&FnSpec> {
        self.funcs.get(name)
    }

    pub fn count(&self) -> usize {
        self.funcs.len()
    }
}

impl Default for FunctionRegistry {
    fn default() -> Self {
        let mut reg = Self::new();
        crate::functions::register_all(&mut reg);
        reg
    }
}
