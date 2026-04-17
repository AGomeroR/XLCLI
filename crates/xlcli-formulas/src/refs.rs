use xlcli_core::types::CellAddr;

use crate::ast::Expr;

pub fn extract_refs(expr: &Expr, current_sheet: u16) -> Vec<CellAddr> {
    let mut refs = Vec::new();
    let no_resolver: Option<&fn(&str) -> Option<u16>> = None;
    collect_refs(expr, current_sheet, &no_resolver, &mut refs);
    refs
}

pub fn extract_refs_with_resolver<F>(expr: &Expr, current_sheet: u16, resolver: F) -> Vec<CellAddr>
where
    F: Fn(&str) -> Option<u16>,
{
    let mut refs = Vec::new();
    collect_refs(expr, current_sheet, &Some(&resolver), &mut refs);
    refs
}

fn collect_refs<F>(expr: &Expr, default_sheet: u16, resolver: &Option<&F>, refs: &mut Vec<CellAddr>)
where
    F: Fn(&str) -> Option<u16>,
{
    match expr {
        Expr::CellRef { sheet: ref s, col, row } => {
            let sheet = resolve_sheet_idx(s, default_sheet, resolver);
            refs.push(CellAddr::new(sheet, row.value(), col.value()));
        }
        Expr::Range { start, end } => {
            if let (
                Expr::CellRef { sheet: ref s, col: c1, row: r1 },
                Expr::CellRef { col: c2, row: r2, .. },
            ) = (start.as_ref(), end.as_ref())
            {
                let sheet = resolve_sheet_idx(s, default_sheet, resolver);
                let r_min = r1.value().min(r2.value());
                let r_max = r1.value().max(r2.value());
                let c_min = c1.value().min(c2.value());
                let c_max = c1.value().max(c2.value());
                for r in r_min..=r_max {
                    for c in c_min..=c_max {
                        refs.push(CellAddr::new(sheet, r, c));
                    }
                }
            }
        }
        Expr::BinOp { left, right, .. } => {
            collect_refs(left, default_sheet, resolver, refs);
            collect_refs(right, default_sheet, resolver, refs);
        }
        Expr::UnaryOp { expr, .. } | Expr::Percent(expr) => {
            collect_refs(expr, default_sheet, resolver, refs);
        }
        Expr::FnCall { args, .. } => {
            for arg in args {
                collect_refs(arg, default_sheet, resolver, refs);
            }
        }
        // NamedRef deps can't be resolved without workbook context — skip for now
        Expr::NamedRef(_) => {}
        Expr::Number(_) | Expr::String(_) | Expr::Boolean(_) | Expr::Error(_) => {}
    }
}

fn resolve_sheet_idx<F>(name: &Option<String>, default: u16, resolver: &Option<&F>) -> u16
where
    F: Fn(&str) -> Option<u16>,
{
    match (name, resolver) {
        (Some(n), Some(f)) => f(n).unwrap_or(default),
        _ => default,
    }
}
