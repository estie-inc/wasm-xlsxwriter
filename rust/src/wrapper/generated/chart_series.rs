use crate::wrapper::ChartDataLabel;
use crate::wrapper::ChartErrorBars;
use crate::wrapper::ChartMarker;
use crate::wrapper::ChartPoint;
use crate::wrapper::ChartTrendline;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartSeries {
    pub(crate) parent: Arc<Mutex<xlsx::Chart>>,
}

#[wasm_bindgen]
impl ChartSeries {
    #[wasm_bindgen(js_name = "setSecondaryAxis", skip_jsdoc)]
    pub fn set_secondary_axis(&self, enable: bool) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().set_secondary_axis(enable);
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setMarker", skip_jsdoc)]
    pub fn set_marker(&self, marker: ChartMarker) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().set_marker(&marker.inner);
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setDataLabel", skip_jsdoc)]
    pub fn set_data_label(&self, data_label: ChartDataLabel) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().set_data_label(&data_label.inner);
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setCustomDataLabels", skip_jsdoc)]
    pub fn set_custom_data_labels(&self, data_labels: Vec<ChartDataLabel>) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().set_custom_data_labels(
            &data_labels
                .iter()
                .map(|x| x.inner.clone())
                .collect::<Vec<_>>(),
        );
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setPoints", skip_jsdoc)]
    pub fn set_points(&self, points: Vec<ChartPoint>) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series()
            .set_points(&points.iter().map(|x| x.inner.clone()).collect::<Vec<_>>());
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setTrendline", skip_jsdoc)]
    pub fn set_trendline(&self, trendline: ChartTrendline) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series()
            .set_trendline(&*trendline.inner.lock().unwrap());
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setYErrorBars", skip_jsdoc)]
    pub fn set_y_error_bars(&self, error_bars: ChartErrorBars) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series()
            .set_y_error_bars(&*error_bars.inner.lock().unwrap());
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setXErrorBars", skip_jsdoc)]
    pub fn set_x_error_bars(&self, error_bars: ChartErrorBars) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series()
            .set_x_error_bars(&*error_bars.inner.lock().unwrap());
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setOverlap", skip_jsdoc)]
    pub fn set_overlap(&self, overlap: i8) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().set_overlap(overlap);
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setGap", skip_jsdoc)]
    pub fn set_gap(&self, gap: u16) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().set_gap(gap);
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setSmooth", skip_jsdoc)]
    pub fn set_smooth(&self, enable: bool) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().set_smooth(enable);
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "setInvertIfNegative", skip_jsdoc)]
    pub fn set_invert_if_negative(&self) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().set_invert_if_negative();
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
    #[wasm_bindgen(js_name = "deleteFromLegend", skip_jsdoc)]
    pub fn delete_from_legend(&self, enable: bool) -> ChartSeries {
        let mut lock = self.parent.lock().unwrap();
        lock.add_series().delete_from_legend(enable);
        ChartSeries {
            parent: Arc::clone(&self.parent),
        }
    }
}
