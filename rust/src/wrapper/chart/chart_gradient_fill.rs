use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;
use crate::wrapper::chart::chart_gradient_fill_type::ChartGradientFillType;
use crate::wrapper::chart::chart_gradient_stop::ChartGradientStop;

#[wasm_bindgen]
pub struct ChartGradientFill {
    pub(crate) inner: xlsx::ChartGradientFill,
}

#[wasm_bindgen]
impl ChartGradientFill {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartGradientFill {
        ChartGradientFill {
            inner: xlsx::ChartGradientFill::new(),
        }
    }

    #[wasm_bindgen(js_name = "setType")]
    pub fn set_type(&mut self, fill_type: ChartGradientFillType) -> ChartGradientFill {
        self.inner.set_type(fill_type.into());
        ChartGradientFill {
            inner: self.inner.clone(),
        }
    }

    #[wasm_bindgen(js_name = "setGradientStops")]
    pub fn set_gradient_stops(&mut self, gradient_stops: Vec<ChartGradientStop>) -> ChartGradientFill {
        let gradient_stops_vec: Vec<xlsx::ChartGradientStop> = gradient_stops.iter().map(|s| s.inner.clone()).collect();
        self.inner.set_gradient_stops(&gradient_stops_vec);
        ChartGradientFill {
            inner: self.inner.clone(),
        }
    }

    #[wasm_bindgen(js_name = "setAngle")]
    pub fn set_angle(&mut self, angle: u16) -> ChartGradientFill {
        self.inner.set_angle(angle);
        ChartGradientFill {
            inner: self.inner.clone(),
        }
    }
}
