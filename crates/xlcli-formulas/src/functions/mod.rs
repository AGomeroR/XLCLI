pub mod math;
pub mod text;
pub mod logical;
pub mod lookup;
pub mod stat;
pub mod info;
pub mod date;

use crate::registry::FunctionRegistry;

pub fn register_all(reg: &mut FunctionRegistry) {
    math::register(reg);
    text::register(reg);
    logical::register(reg);
    lookup::register(reg);
    stat::register(reg);
    info::register(reg);
    date::register(reg);
}
