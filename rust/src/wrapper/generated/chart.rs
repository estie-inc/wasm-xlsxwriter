use crate::wrapper::ChartDataTable;
use crate::wrapper::ChartSeries;
use crate::wrapper::ChartType;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct Chart {
    pub(crate) inner: Arc<Mutex<xlsx::Chart>>,
}

#[wasm_bindgen]
impl Chart {
    #[wasm_bindgen(constructor)]
    pub fn new(chart_type: ChartType) -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new(chart_type.inner))),
        }
    }
    #[wasm_bindgen(js_name = "pushSeries", skip_jsdoc)]
    pub fn push_series(&self, series: ChartSeries) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.push_series(&series.inner);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "title", skip_jsdoc)]
    pub fn title(&self) -> &mut ChartTitle {
        let lock = self.inner.lock().unwrap();
        lock.title()
    }
    #[wasm_bindgen(js_name = "xAxis", skip_jsdoc)]
    pub fn x_axis(&self) -> &mut ChartAxis {
        let lock = self.inner.lock().unwrap();
        lock.x_axis()
    }
    #[wasm_bindgen(js_name = "yAxis", skip_jsdoc)]
    pub fn y_axis(&self) -> &mut ChartAxis {
        let lock = self.inner.lock().unwrap();
        lock.y_axis()
    }
    #[wasm_bindgen(js_name = "x2Axis", skip_jsdoc)]
    pub fn x2_axis(&self) -> &mut ChartAxis {
        let lock = self.inner.lock().unwrap();
        lock.x2_axis()
    }
    #[wasm_bindgen(js_name = "y2Axis", skip_jsdoc)]
    pub fn y2_axis(&self) -> &mut ChartAxis {
        let lock = self.inner.lock().unwrap();
        lock.y2_axis()
    }
    #[wasm_bindgen(js_name = "legend", skip_jsdoc)]
    pub fn legend(&self) -> &mut ChartLegend {
        let lock = self.inner.lock().unwrap();
        lock.legend()
    }
    #[wasm_bindgen(js_name = "chartArea", skip_jsdoc)]
    pub fn chart_area(&self) -> &mut ChartArea {
        let lock = self.inner.lock().unwrap();
        lock.chart_area()
    }
    #[wasm_bindgen(js_name = "plotArea", skip_jsdoc)]
    pub fn plot_area(&self) -> &mut ChartPlotArea {
        let lock = self.inner.lock().unwrap();
        lock.plot_area()
    }
    #[wasm_bindgen(js_name = "combine", skip_jsdoc)]
    pub fn combine(&self, chart: Chart) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.combine(&chart.inner);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setStyle", skip_jsdoc)]
    pub fn set_style(&self, style: u8) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_style(style);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setRotation", skip_jsdoc)]
    pub fn set_rotation(&self, rotation: u16) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_rotation(rotation);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setHoleSize", skip_jsdoc)]
    pub fn set_hole_size(&self, hole_size: u8) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_hole_size(hole_size);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setUpDownBars", skip_jsdoc)]
    pub fn set_up_down_bars(&self, enable: bool) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_up_down_bars(enable);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setHighLowLines", skip_jsdoc)]
    pub fn set_high_low_lines(&self, enable: bool) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_high_low_lines(enable);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setDropLines", skip_jsdoc)]
    pub fn set_drop_lines(&self, enable: bool) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_drop_lines(enable);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setDataTable", skip_jsdoc)]
    pub fn set_data_table(&self, table: ChartDataTable) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_data_table(&*table.inner.lock().unwrap());
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setWidth", skip_jsdoc)]
    pub fn set_width(&self, width: u32) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_width(width);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setHeight", skip_jsdoc)]
    pub fn set_height(&self, height: u32) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_height(height);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setScaleHeight", skip_jsdoc)]
    pub fn set_scale_height(&self, scale: f64) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_scale_height(scale);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setScaleWidth", skip_jsdoc)]
    pub fn set_scale_width(&self, scale: f64) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_scale_width(scale);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, name: &str) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_name(name);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setAltText", skip_jsdoc)]
    pub fn set_alt_text(&self, alt_text: &str) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_alt_text(alt_text);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "setDecorative", skip_jsdoc)]
    pub fn set_decorative(&self, enable: bool) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_decorative(enable);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "showNaAsEmptyCell", skip_jsdoc)]
    pub fn show_na_as_empty_cell(&self) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.show_na_as_empty_cell();
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    #[wasm_bindgen(js_name = "showHiddenData", skip_jsdoc)]
    pub fn show_hidden_data(&self) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.show_hidden_data();
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
}
