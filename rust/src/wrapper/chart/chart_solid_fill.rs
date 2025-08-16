use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

use crate::wrapper::color::Color;

/// The `ChartSolidFill` struct represents a solid fill for chart elements.
///
/// The `ChartSolidFill` struct is used to define the solid fill properties
/// for chart elements such as plot areas, chart areas, data series, and other
/// fillable elements in a chart.
///
/// It is used in conjunction with the {@link Chart} struct.
#[wasm_bindgen]
#[derive(Clone)]
pub struct ChartSolidFill {
    pub(crate) inner: xlsx::ChartSolidFill,
}

#[wasm_bindgen]
impl ChartSolidFill {
    /// Create a new `ChartSolidFill` object.
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartSolidFill {
        ChartSolidFill {
            inner: xlsx::ChartSolidFill::new(),
        }
    }

    /// Set the color of a solid fill.
    ///
    /// @param {Color} color - The color property.
    /// @return {ChartSolidFill} - The ChartSolidFill instance.
    #[wasm_bindgen(js_name = "setColor")]
    pub fn set_color(&mut self, color: Color) -> ChartSolidFill {
        self.inner.set_color(color);
        self.clone()
    }

    /// Set the transparency of a solid fill.
    ///
    /// Set the transparency of a solid fill color for a Chart element.
    /// You must also specify a fill color in order for the transparency to be applied.
    ///
    /// @param {number} transparency - The color transparency in the range 0-100.
    /// @return {ChartSolidFill} - The ChartSolidFill instance.
    #[wasm_bindgen(js_name = "setTransparency")]
    pub fn set_transparency(&mut self, transparency: u8) -> ChartSolidFill {
        self.inner.set_transparency(transparency);
        self.clone()
    }
}
