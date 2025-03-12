mod chart_axis;
mod chart_font;
mod chart_format;
mod chart_legend;
mod chart_legend_position;
mod chart_point;
mod chart_range;
mod chart_series;
mod chart_title;
mod chart_type;

use std::sync::{Arc, Mutex};

use chart_axis::ChartAxis;
use chart_legend::ChartLegend;
use chart_series::ChartSeries;
use chart_title::ChartTitle;
use chart_type::ChartType;
use rust_xlsxwriter as xlsx;
use wasm_bindgen::prelude::*;

#[wasm_bindgen]
pub struct Chart {
    pub(crate) inner: Arc<Mutex<xlsx::Chart>>,
}

#[wasm_bindgen]
impl Chart {
    /// Create a new Chart object.
    ///
    /// Create a new Chart object that can be inserted into a worksheet.
    ///
    /// @param {u8} chart_type - The type of the chart.
    /// @returns {Chart} - The Chart object.
    #[wasm_bindgen(constructor)]
    pub fn new(chart_type: ChartType) -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new(chart_type.into()))),
        }
    }

    pub(crate) fn lock(&self) -> std::sync::MutexGuard<'_, xlsx::Chart> {
        self.inner.lock().unwrap()
    }

    #[wasm_bindgen(js_name = "newArea")]
    pub fn new_area() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_area())),
        }
    }

    #[wasm_bindgen(js_name = "newBar")]
    pub fn new_bar() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_bar())),
        }
    }

    #[wasm_bindgen(js_name = "newColumn")]
    pub fn new_column() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_column())),
        }
    }

    #[wasm_bindgen(js_name = "newColumnStacked")]
    pub fn new_doughnut() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_doughnut())),
        }
    }

    #[wasm_bindgen(js_name = "newLine")]
    pub fn new_line() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_line())),
        }
    }

    #[wasm_bindgen(js_name = "newPie")]
    pub fn new_pie() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_pie())),
        }
    }

    #[wasm_bindgen(js_name = "newRadar")]
    pub fn new_radar() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_radar())),
        }
    }

    #[wasm_bindgen(js_name = "newScatter")]
    pub fn new_scatter() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_scatter())),
        }
    }

    #[wasm_bindgen(js_name = "newStock")]
    pub fn new_stock() -> Chart {
        Chart {
            inner: Arc::new(Mutex::new(xlsx::Chart::new_stock())),
        }
    }

    // FIXME: add_series not supported for ownership reasons

    #[wasm_bindgen(js_name = "pushSeries")]
    pub fn push_series(&self, series: &ChartSeries) -> Chart {
        let mut chart = self.inner.lock().unwrap();
        let series = series.inner.lock().unwrap();
        chart.push_series(&series);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }

    #[wasm_bindgen(js_name = "title")]
    pub fn title(&self) -> ChartTitle {
        ChartTitle {
            chart: Arc::clone(&self.inner),
        }
    }

    /// Set a user defined name for a chart.
    ///
    /// By default Excel names charts as "Chart 1", "Chart 2", etc. This name
    /// shows up in the formula bar and can be used to find or reference a
    /// chart.
    ///
    /// The set_name() method allows you to give the chart a user
    /// defined name.
    ///
    /// @param {string} name - A user defined name for the chart.
    /// @returns {Chart} - The Chart object.
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, name: &str) -> Chart {
        let mut chart = self.inner.lock().unwrap();
        chart.set_name(name);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }

    #[wasm_bindgen(js_name = "setAltText", skip_jsdoc)]
    pub fn set_alt_text(&self, alt_text: &str) -> Chart {
        let mut chart = self.inner.lock().unwrap();
        chart.set_alt_text(alt_text);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }

    #[wasm_bindgen(js_name = "setWidth", skip_jsdoc)]
    pub fn set_width(&self, width: u32) -> Chart {
        let mut chart = self.inner.lock().unwrap();
        chart.set_width(width);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }

    #[wasm_bindgen(js_name = "setHeight", skip_jsdoc)]
    pub fn set_height(&self, height: u32) -> Chart {
        let mut chart = self.inner.lock().unwrap();
        chart.set_height(height);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }

    #[wasm_bindgen(js_name = "xAxis", skip_jsdoc)]
    pub fn x_axis(&self) -> ChartAxis {
        ChartAxis {
            inner: Arc::clone(&self.inner),
            axis: chart_axis::AxisType::X,
        }
    }

    #[wasm_bindgen(js_name = "yAxis", skip_jsdoc)]
    pub fn y_axis(&self) -> ChartAxis {
        ChartAxis {
            inner: Arc::clone(&self.inner),
            axis: chart_axis::AxisType::Y,
        }
    }

    #[wasm_bindgen(js_name = "x2Axis", skip_jsdoc)]
    pub fn x2_axis(&self) -> ChartAxis {
        ChartAxis {
            inner: Arc::clone(&self.inner),
            axis: chart_axis::AxisType::X2,
        }
    }

    #[wasm_bindgen(js_name = "y2Axis", skip_jsdoc)]
    pub fn y2_axis(&self) -> ChartAxis {
        ChartAxis {
            inner: Arc::clone(&self.inner),
            axis: chart_axis::AxisType::Y2,
        }
    }

    #[wasm_bindgen(js_name = "legend", skip_jsdoc)]
    pub fn legend(&self) -> ChartLegend {
        ChartLegend {
            chart: Arc::clone(&self.inner),
        }
    }
}
