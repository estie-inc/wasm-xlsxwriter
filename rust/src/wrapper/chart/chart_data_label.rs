use rust_xlsxwriter::{self as xlsx};
use wasm_bindgen::prelude::*;

use crate::wrapper::chart::chart_data_label_position::ChartDataLabelPosition;
use crate::wrapper::chart::chart_font::ChartFont;
use crate::wrapper::chart::chart_format::ChartFormat;

#[wasm_bindgen]
#[derive(Clone)]
pub struct ChartDataLabel {
    pub(crate) inner: xlsx::ChartDataLabel,
}

#[wasm_bindgen]
impl ChartDataLabel {
    #[wasm_bindgen(constructor)]
    pub fn new() -> ChartDataLabel {
        ChartDataLabel {
            inner: xlsx::ChartDataLabel::new(),
        }
    }

    #[wasm_bindgen(js_name = "showValue")]
    pub fn show_value(&mut self) -> ChartDataLabel {
        self.inner.show_value();
        self.clone()
    }

    #[wasm_bindgen(js_name = "showCategoryName")]
    pub fn show_category_name(&mut self) -> ChartDataLabel {
        self.inner.show_category_name();
        self.clone()
    }

    #[wasm_bindgen(js_name = "showSeriesName")]
    pub fn show_series_name(&mut self) -> ChartDataLabel {
        self.inner.show_series_name();
        self.clone()
    }

    #[wasm_bindgen(js_name = "showLeaderLines")]
    pub fn show_leader_lines(&mut self) -> ChartDataLabel {
        self.inner.show_leader_lines();
        self.clone()
    }

    #[wasm_bindgen(js_name = "showLegendKey")]
    pub fn show_legend_key(&mut self) -> ChartDataLabel {
        self.inner.show_legend_key();
        self.clone()
    }

    #[wasm_bindgen(js_name = "showPercentage")]
    pub fn show_percentage(&mut self) -> ChartDataLabel {
        self.inner.show_percentage();
        self.clone()
    }

    #[wasm_bindgen(js_name = "setPosition")]
    pub fn set_position(&mut self, position: ChartDataLabelPosition) -> ChartDataLabel {
        self.inner.set_position(position.into());
        self.clone()
    }

    #[wasm_bindgen(js_name = "setFont")]
    pub fn set_font(&mut self, font: &ChartFont) -> ChartDataLabel {
        self.inner.set_font(&font.inner);
        self.clone()
    }

    #[wasm_bindgen(js_name = "setFormat")]
    pub fn set_format(&mut self, format: &mut ChartFormat) -> ChartDataLabel {
        self.inner.set_format(&mut format.inner);
        self.clone()
    }

    #[wasm_bindgen(js_name = "setNumFormat")]
    pub fn set_num_format(&mut self, num_format: &str) -> ChartDataLabel {
        self.inner.set_num_format(num_format);
        self.clone()
    }

    #[wasm_bindgen(js_name = "setSeparator")]
    pub fn set_separator(&mut self, separator: char) -> ChartDataLabel {
        self.inner.set_separator(separator);
        self.clone()
    }

    #[wasm_bindgen(js_name = "showYValue")]
    pub fn show_y_value(&mut self) -> ChartDataLabel {
        self.inner.show_y_value();
        self.clone()
    }

    #[wasm_bindgen(js_name = "showXValue")]
    pub fn show_x_value(&mut self) -> ChartDataLabel {
        self.inner.show_x_value();
        self.clone()
    }

    #[wasm_bindgen(js_name = "setHidden")]
    pub fn set_hidden(&mut self) -> ChartDataLabel {
        self.inner.set_hidden();
        self.clone()
    }

    #[wasm_bindgen(js_name = "setValue")]
    pub fn set_value(&mut self, value: &str) -> ChartDataLabel {
        self.inner.set_value(value);
        self.clone()
    }

    #[wasm_bindgen(js_name = "toCustom")]
    pub fn to_custom(&mut self) -> ChartDataLabel {
        self.inner.to_custom();
        self.clone()
    }
}

