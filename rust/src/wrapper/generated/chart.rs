use crate::wrapper::ChartDataTable;
use crate::wrapper::ChartSeries;
use crate::wrapper::ChartType;
use crate::wrapper::ObjectMovement;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

/// The `Chart` struct is used to create an object to represent an chart that
/// can be inserted into a worksheet.
///
/// The `Chart` struct exposes other chart related structs that allow you to
/// configure the chart such as {@link ChartSeries}, {@link ChartAxis} and
/// {@link ChartTitle}.
///
/// Charts are added to the worksheets using the
/// {@link Worksheet#insertChart} or
/// {@link Worksheet#insertChartWithOffset}
/// methods.
///
/// See also Working with Charts for a general introduction to
/// working with charts in `rust_xlsxwriter`.
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
    /// Add a chart series to a chart.
    ///
    /// Add a standalone chart series to a chart. The chart series represents
    /// the category and value ranges as well as formatting and display options.
    /// A chart in Excel must contain at least one data series. A series is
    /// represented by a {@link ChartSeries} struct.
    ///
    /// A chart series is usually created via the
    /// {@link Chart#addSeries} method, see above. However, if
    /// required you can create a standalone `ChartSeries` object and add it to
    /// a chart via this `chart.push_series()` method.
    ///
    /// # Parameters
    ///
    /// `series` - a {@link ChartSeries} instance.
    #[wasm_bindgen(js_name = "pushSeries", skip_jsdoc)]
    pub fn push_series(&self, series: ChartSeries) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.push_series(&*series.inner.lock().unwrap());
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Get the chart title object in order to set its properties.
    ///
    /// Get a reference to the chart's X-Axis {@link ChartTitle} object in order to
    /// set its properties.
    #[wasm_bindgen(js_name = "title", skip_jsdoc)]
    pub fn title(&self) -> &mut ChartTitle {
        let lock = self.inner.lock().unwrap();
        lock.title()
    }
    /// Get the chart X-Axis object in order to set its properties.
    ///
    /// Get a reference to the chart's X-Axis {@link ChartAxis} object in order to
    /// set its properties.
    #[wasm_bindgen(js_name = "xAxis", skip_jsdoc)]
    pub fn x_axis(&self) -> &mut ChartAxis {
        let lock = self.inner.lock().unwrap();
        lock.x_axis()
    }
    /// Get the chart Y-Axis object in order to set its properties.
    ///
    /// Get a reference to the chart's Y-Axis {@link ChartAxis} object in order to
    /// set its properties.
    ///
    /// See the {@link Chart#xAxis} method above.
    #[wasm_bindgen(js_name = "yAxis", skip_jsdoc)]
    pub fn y_axis(&self) -> &mut ChartAxis {
        let lock = self.inner.lock().unwrap();
        lock.y_axis()
    }
    /// Get the chart X2-Axis object in order to set its properties.
    ///
    /// Get a reference to the chart's X2-Axis {@link ChartAxis} object in order to
    /// set its properties.
    ///
    /// See the {@link Chart#xAxis} method above.
    #[wasm_bindgen(js_name = "x2Axis", skip_jsdoc)]
    pub fn x2_axis(&self) -> &mut ChartAxis {
        let lock = self.inner.lock().unwrap();
        lock.x2_axis()
    }
    /// Get the chart Y2-Axis object in order to set its properties.
    ///
    /// Get a reference to the chart's Y2-Axis {@link ChartAxis} object in order to
    /// set its properties.
    ///
    /// See the {@link Chart#xAxis} method above.
    #[wasm_bindgen(js_name = "y2Axis", skip_jsdoc)]
    pub fn y2_axis(&self) -> &mut ChartAxis {
        let lock = self.inner.lock().unwrap();
        lock.y2_axis()
    }
    /// Get the chart legend object in order to set its properties.
    ///
    /// Get a reference to the chart's {@link ChartLegend} object in order to set
    /// its properties.
    #[wasm_bindgen(js_name = "legend", skip_jsdoc)]
    pub fn legend(&self) -> &mut ChartLegend {
        let lock = self.inner.lock().unwrap();
        lock.legend()
    }
    /// Get the chart area object in order to set its properties.
    ///
    /// Get a reference to the chart's {@link ChartArea} object in order to set its
    /// properties. The `ChartArea` is a representation of the background area
    /// object of an Excel chart.
    #[wasm_bindgen(js_name = "chartArea", skip_jsdoc)]
    pub fn chart_area(&self) -> &mut ChartArea {
        let lock = self.inner.lock().unwrap();
        lock.chart_area()
    }
    /// Get the chart plot area object in order to set its properties.
    ///
    /// Get a reference to the chart's {@link ChartPlotArea} object in order to set
    /// its properties. The `ChartPlotArea` struct is a representation of the
    /// plotting area an Excel chart.
    #[wasm_bindgen(js_name = "plotArea", skip_jsdoc)]
    pub fn plot_area(&self) -> &mut ChartPlotArea {
        let lock = self.inner.lock().unwrap();
        lock.plot_area()
    }
    /// Create a combined chart from two different chart types.
    ///
    /// In Excel is also possible to combine two different chart types, for
    /// example a column and line chart to create a Pareto chart. In
    /// `rust_xlsxwriter` you can replicate this by creating a new chart
    /// instance as the primary chart and then create a secondary chart of a
    /// different type and combine it with the primary chart using the
    /// `Chart::combine()` method.
    ///
    /// The combined secondary chart can share the same Y axis as the primary
    /// chart or it can use a secondary Y2 axis. An example of each is shown
    /// below.
    ///
    /// See Combined Charts for additional
    /// information on combined charts and also some limitations.
    ///
    /// # Parameters
    ///
    /// - `chart`: The {@link Chart} to insert into the cell.
    #[wasm_bindgen(js_name = "combine", skip_jsdoc)]
    pub fn combine(&self, chart: Chart) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.combine(&*chart.inner.lock().unwrap());
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the chart style type.
    ///
    /// The `set_style()` method is used to set the style of the chart to one of
    /// 48 built-in styles.
    ///
    /// These styles were available in the original Excel 2007 interface. In
    /// later versions they have been replaced with "layouts" on the "Chart
    /// Design" tab. These layouts are not defined in the file format. They are
    /// a collection of modifications to the base chart type. They can be
    /// replicated using the Chart APIs but they cannot be defined by the
    /// `set_style()` method.
    ///
    /// # Parameters
    ///
    /// - `style`: A integer value in the range 1-48.
    #[wasm_bindgen(js_name = "setStyle", skip_jsdoc)]
    pub fn set_style(&self, style: u8) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_style(style);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the Pie/Doughnut chart rotation.
    ///
    /// The `set_rotation()` method is used to set the rotation of the first
    /// segment of a Pie/Doughnut chart. This has the effect of rotating the
    /// entire chart.
    ///
    /// # Parameters
    ///
    /// - `rotation`: The rotation of the first segment of a Pie/Doughnut chart.
    ///   The range is 0 <= rotation <= 360 and the default is 0.
    #[wasm_bindgen(js_name = "setRotation", skip_jsdoc)]
    pub fn set_rotation(&self, rotation: u16) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_rotation(rotation);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the hole size for a Doughnut chart.
    ///
    /// Set the center hole size for a Doughnut chart.
    ///
    /// # Parameters
    ///
    /// - `hole_size`: The hole size for a Doughnut chart. The range is 0 <=
    ///   `hole_size` <= 90 and the default is 50.
    #[wasm_bindgen(js_name = "setHoleSize", skip_jsdoc)]
    pub fn set_hole_size(&self, hole_size: u8) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_hole_size(hole_size);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set Up-Down bar indicators for a Line chart.
    ///
    /// Set Up-Down bar indicator to indicate change between two or more series.
    /// In Excel these can only be added to Line and Stock charts.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setUpDownBars", skip_jsdoc)]
    pub fn set_up_down_bars(&self, enable: bool) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_up_down_bars(enable);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set High-Low lines for a Line chart.
    ///
    /// Set High-Low lines for a Line chart to indicate the high and low values
    /// between two or more series. In Excel these can only be added to Line and
    /// Stock charts.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setHighLowLines", skip_jsdoc)]
    pub fn set_high_low_lines(&self, enable: bool) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_high_low_lines(enable);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set drop lines for a chart.
    ///
    /// Set drop lines for a chart between the maximum value and the associated
    /// category value. In Excel these can only be added to Line, Area and Stock
    /// charts.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setDropLines", skip_jsdoc)]
    pub fn set_drop_lines(&self, enable: bool) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_drop_lines(enable);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set a data table for a chart.
    ///
    /// A chart data table in Excel is an additional table below a chart that
    /// shows the plotted data in tabular form.
    ///
    /// The chart data table has the following default properties which can be
    /// set via the {@link ChartDataTable} struct.
    ///
    /// src="https://rustxlsxwriter.github.io/images/chart_data_table_options.png">
    ///
    /// # Parameters
    ///
    /// - `table`: A {@link ChartDataTable} reference.
    #[wasm_bindgen(js_name = "setDataTable", skip_jsdoc)]
    pub fn set_data_table(&self, table: ChartDataTable) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_data_table(&*table.inner.lock().unwrap());
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the width of the chart.
    ///
    /// The default width of an Excel chart is 480 pixels. The `set_width()`
    /// method allows you to set it to some other non-zero size.
    ///
    /// # Parameters
    ///
    /// - `width`: The chart width in pixels.
    #[wasm_bindgen(js_name = "setWidth", skip_jsdoc)]
    pub fn set_width(&self, width: u32) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_width(width);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the height of the chart.
    ///
    /// The default height of an Excel chart is 480 pixels. The `set_height()`
    /// method allows you to set it to some other non-zero size. See the example
    /// above.
    ///
    /// # Parameters
    ///
    /// - `height`: The chart height in pixels.
    #[wasm_bindgen(js_name = "setHeight", skip_jsdoc)]
    pub fn set_height(&self, height: u32) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_height(height);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the height scale for the chart.
    ///
    /// Set the height scale for the chart relative to 1.0 (i.e. 100%). This is
    /// a syntactic alternative to {@link Chart#setHeight}.
    ///
    /// # Parameters
    ///
    /// - `scale`: The scale ratio.
    #[wasm_bindgen(js_name = "setScaleHeight", skip_jsdoc)]
    pub fn set_scale_height(&self, scale: f64) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_scale_height(scale);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the width scale for the chart.
    ///
    /// Set the width scale for the chart relative to 1.0 (i.e. 100%). This is a
    /// syntactic alternative to {@link Chart#setWidth}.
    ///
    /// # Parameters
    ///
    /// - `scale`: The scale ratio.
    #[wasm_bindgen(js_name = "setScaleWidth", skip_jsdoc)]
    pub fn set_scale_width(&self, scale: f64) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_scale_width(scale);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set a user defined name for a chart.
    ///
    /// By default Excel names charts as "Chart 1", "Chart 2", etc. This name
    /// shows up in the formula bar and can be used to find or reference a
    /// chart.
    ///
    /// The {@link Chart#setName} method allows you to give the chart a user
    /// defined name.
    ///
    /// # Parameters
    ///
    /// - `name`: A user defined name for the chart.
    #[wasm_bindgen(js_name = "setName", skip_jsdoc)]
    pub fn set_name(&self, name: &str) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_name(name);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the alt text for the chart.
    ///
    /// Set the alt text for the chart to help accessibility. The alt text is
    /// used with screen readers to help people with visual disabilities.
    ///
    /// See the following Microsoft documentation on [Everything you need to
    /// know to write effective alt
    /// text](https://support.microsoft.com/en-us/office/everything-you-need-to-know-to-write-effective-alt-text-df98f884-ca3d-456c-807b-1a1fa82f5dc2).
    ///
    /// # Parameters
    ///
    /// - `alt_text`: The alt text string to add to the chart.
    #[wasm_bindgen(js_name = "setAltText", skip_jsdoc)]
    pub fn set_alt_text(&self, alt_text: &str) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_alt_text(alt_text);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Mark a chart as decorative.
    ///
    /// Charts don't always need an alt text description. Some charts may contain
    /// little or no useful visual information. Such charts can be marked as
    /// "decorative" so that screen readers can inform the users that they don't
    /// contain important information.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setDecorative", skip_jsdoc)]
    pub fn set_decorative(&self, enable: bool) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_decorative(enable);
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set the object movement options for a chart.
    ///
    /// Set the option to define how an chart will behave in Excel if the cells
    /// under the chart are moved, deleted, or have their size changed. In Excel
    /// the options are:
    ///
    /// 1. Move and size with cells. Default for charts.
    /// 2. Move but don't size with cells.
    /// 3. Don't move or size with cells.
    ///
    /// These values are defined in the {@link ObjectMovement} enum.
    ///
    /// # Parameters
    ///
    /// `option` - A {@link ObjectMovement} enum value.
    #[wasm_bindgen(js_name = "setObjectMovement", skip_jsdoc)]
    pub fn set_object_movement(&self, option: ObjectMovement) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.set_object_movement(xlsx::ObjectMovement::from(option));
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display #N/A on charts as blank/empty cells.
    #[wasm_bindgen(js_name = "showNaAsEmptyCell", skip_jsdoc)]
    pub fn show_na_as_empty_cell(&self) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.show_na_as_empty_cell();
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Display data on charts from hidden rows or columns.
    #[wasm_bindgen(js_name = "showHiddenData", skip_jsdoc)]
    pub fn show_hidden_data(&self) -> Chart {
        let mut lock = self.inner.lock().unwrap();
        lock.show_hidden_data();
        Chart {
            inner: Arc::clone(&self.inner),
        }
    }
    /// Set default values for the primary chart axis ids.
    ///
    /// This is mainly used to ensure that the primary axis ids used in testing
    /// match the semi-randomized values in the target Excel file.
    ///
    /// # Parameters
    ///
    /// - `axis_id1`: X-axis id.
    /// - `axis_id2`: Y-axis id.
    #[wasm_bindgen(js_name = "setAxisIds", skip_jsdoc)]
    pub fn set_axis_ids(&self, axis_id1: u32, axis_id2: u32) -> () {
        let mut lock = self.inner.lock().unwrap();
        lock.set_axis_ids(axis_id1, axis_id2);
    }
    /// Set default values for the secondary chart axis ids.
    ///
    /// This is mainly used to ensure that the secondary axis ids used in
    /// testing match the semi-randomized values in the target Excel file.
    ///
    /// # Parameters
    ///
    /// - `axis_id1`: X-axis id.
    /// - `axis_id2`: Y-axis id.
    #[wasm_bindgen(js_name = "setAxis2Ids", skip_jsdoc)]
    pub fn set_axis2_ids(&self, axis_id1: u32, axis_id2: u32) -> () {
        let mut lock = self.inner.lock().unwrap();
        lock.set_axis2_ids(axis_id1, axis_id2);
    }
}
