use crate::wrapper::ChartAxisCrossing;
use crate::wrapper::ChartAxisDateUnitType;
use crate::wrapper::ChartAxisDisplayUnitType;
use crate::wrapper::ChartAxisLabelAlignment;
use crate::wrapper::ChartAxisLabelPosition;
use crate::wrapper::ChartAxisTickType;
use crate::wrapper::ChartFont;
use crate::wrapper::ChartLayout;
use crate::wrapper::ChartLine;
use crate::wrapper::WasmResult;
use rust_xlsxwriter as xlsx;
use std::sync::{Arc, Mutex};
use wasm_bindgen::prelude::*;

#[derive(Clone, Copy)]
pub enum ChartAxisAccessor {
    XAxis,
    YAxis,
    X2Axis,
    Y2Axis,
}

/// The `ChartAxis` struct represents a chart axis.
///
/// Used in conjunction with the {@link Chart#xAxis} and {@link Chart#yAxis}.
///
/// It is used in conjunction with the {@link Chart} struct.
#[derive(Clone)]
#[wasm_bindgen]
pub struct ChartAxis {
    pub(crate) parent: Arc<Mutex<xlsx::Chart>>,
    pub(crate) accessor: ChartAxisAccessor,
}

#[wasm_bindgen]
impl ChartAxis {
    /// Set the font properties of a chart axis title.
    ///
    /// Set the font properties of a chart axis name/title using a {@link ChartFont}
    /// reference. Example font properties that can be set are:
    ///
    /// - {@link ChartFont#setBold}
    /// - {@link ChartFont#setItalic}
    /// - {@link ChartFont#setName}
    /// - {@link ChartFont#setSize}
    /// - {@link ChartFont#setRotation}
    ///
    /// See {@link ChartFont} for full details.
    ///
    /// The name font property for an axis represents the font for the axis
    /// title. To set the font for the category or value numbers use the
    /// {@link ChartAxis#setFont} method.
    ///
    /// # Parameters
    ///
    /// `font`: A {@link ChartFont} struct reference to represent the font
    /// properties.
    #[wasm_bindgen(js_name = "setNameFont", skip_jsdoc)]
    pub fn set_name_font(&self, font: ChartFont) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_name_font(&font.inner),
            ChartAxisAccessor::YAxis => lock.y_axis().set_name_font(&font.inner),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_name_font(&font.inner),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_name_font(&font.inner),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the font properties of a chart axis.
    ///
    /// Set the font properties of a chart axis using a {@link ChartFont} reference.
    /// Example font properties that can be set are:
    ///
    /// - {@link ChartFont#setBold}
    /// - {@link ChartFont#setItalic}
    /// - {@link ChartFont#setName}
    /// - {@link ChartFont#setSize}
    /// - {@link ChartFont#setRotation}
    ///
    /// See {@link ChartFont} for full details.
    ///
    /// The font property for an axis represents the font for the category or
    /// value names or numbers. To set the font for the axis name/title use the
    /// {@link ChartAxis#setNameFont} method.
    ///
    /// # Parameters
    ///
    /// `font`: A {@link ChartFont} struct reference to represent the font
    /// properties.
    #[wasm_bindgen(js_name = "setFont", skip_jsdoc)]
    pub fn set_font(&self, font: ChartFont) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_font(&font.inner),
            ChartAxisAccessor::YAxis => lock.y_axis().set_font(&font.inner),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_font(&font.inner),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_font(&font.inner),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the number format for a chart axis.
    ///
    /// Excel plots/displays data in charts with the same number formatting that
    /// it has in the worksheet. The `set_num_format()` method allows you to
    /// override this and controls whether a number is displayed as an integer,
    /// a floating point number, a date, a currency value or some other user
    /// defined format.
    ///
    /// See also [Number Format Categories] and [Number Formats in different
    /// locales] in the documentation for {@link Format}.
    ///
    /// [Number Format Categories]: crate::Format#number-format-categories
    /// [Number Formats in different locales]:
    ///     crate::Format#number-formats-in-different-locales
    ///
    /// # Parameters
    ///
    /// - `num_format`: The number format property.
    #[wasm_bindgen(js_name = "setNumFormat", skip_jsdoc)]
    pub fn set_num_format(&self, num_format: &str) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_num_format(num_format),
            ChartAxisAccessor::YAxis => lock.y_axis().set_num_format(num_format),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_num_format(num_format),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_num_format(num_format),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the category axis as a date axis.
    ///
    /// In general the "Category" axis (usually the X-axis) in Excel charts is
    /// made up of evenly spaced categories. This type of axis doesn't support
    /// features such as maximum and minimum even if the categories are numbers.
    /// The two exceptions to this are the "Value" axes used in Scatter charts
    /// and "Date" axes. Date axes are a combination of "Category" and "Value"
    /// axes and they support features of both types of axes.
    ///
    /// In order to have a date axes in your chart you need to have a range of
    /// Date/Time values in a worksheet that the
    /// {@link ChartSeries#setCategories} refer to. You can then use the
    /// `set_date_axis()` method turns on the "date axis" property for a chart
    /// axis.
    ///
    /// See [Chart Value and Category Axes] for an explanation of the
    /// difference between Value and Category axes in Excel.
    ///
    /// [Chart Value and Category Axes]:
    ///     crate::chart#chart-value-and-category-axes
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setDateAxis", skip_jsdoc)]
    pub fn set_date_axis(&self, enable: bool) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_date_axis(enable),
            ChartAxisAccessor::YAxis => lock.y_axis().set_date_axis(enable),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_date_axis(enable),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_date_axis(enable),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the category axis as a text axis.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setTextAxis", skip_jsdoc)]
    pub fn set_text_axis(&self, enable: bool) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_text_axis(enable),
            ChartAxisAccessor::YAxis => lock.y_axis().set_text_axis(enable),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_text_axis(enable),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_text_axis(enable),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the category axis as an automatic axis - generally the default.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setAutomaticAxis", skip_jsdoc)]
    pub fn set_automatic_axis(&self, enable: bool) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_automatic_axis(enable),
            ChartAxisAccessor::YAxis => lock.y_axis().set_automatic_axis(enable),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_automatic_axis(enable),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_automatic_axis(enable),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the crossing point for the opposite axis.
    ///
    /// By default Excel sets chart axes to cross at 0. If required you can use
    /// {@link ChartAxis#setCrossing} and {@link ChartAxisCrossing} to define
    /// another point where the opposite axis will cross the current axis.
    ///
    /// The {@link ChartAxisCrossing} enum defines values like `max` and `min` but
    /// also allows you to define a category value for X-axes (except for
    /// Scatter and Date axes) and an actual value for Y-axes and Scatter and
    /// Date axes.
    ///
    /// # Parameters
    ///
    /// - `crossing`: A {@link ChartAxisCrossing} enum value.
    #[wasm_bindgen(js_name = "setCrossing", skip_jsdoc)]
    pub fn set_crossing(&self, crossing: ChartAxisCrossing) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock
                .x_axis()
                .set_crossing(xlsx::ChartAxisCrossing::from(crossing)),
            ChartAxisAccessor::YAxis => lock
                .y_axis()
                .set_crossing(xlsx::ChartAxisCrossing::from(crossing)),
            ChartAxisAccessor::X2Axis => lock
                .x2_axis()
                .set_crossing(xlsx::ChartAxisCrossing::from(crossing)),
            ChartAxisAccessor::Y2Axis => lock
                .y2_axis()
                .set_crossing(xlsx::ChartAxisCrossing::from(crossing)),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Reverse the direction of the axis categories or values.
    ///
    /// Reverse the direction that the axis data is plotted in from left to
    /// right or top to bottom.
    #[wasm_bindgen(js_name = "setReverse", skip_jsdoc)]
    pub fn set_reverse(&self) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_reverse(),
            ChartAxisAccessor::YAxis => lock.y_axis().set_reverse(),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_reverse(),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_reverse(),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the maximum value for an axis.
    ///
    /// Set the maximum bound to be displayed for an axis.
    ///
    /// Maximum and minimum chart axis values can only be set for chart "Value"
    /// axes and "Category Date" axes in Excel. You cannot set a maximum or
    /// minimum value for "Category" axes even if the category values are
    /// numbers. See [Chart Value and Category Axes] for an explanation of the
    /// difference between Value and Category axes in Excel.
    ///
    /// See also {@link ChartAxis#setMaxDate} below.
    ///
    /// [Chart Value and Category Axes]:
    ///     crate::chart#chart-value-and-category-axes
    ///
    /// # Parameters
    ///
    /// `max` - The maximum bound for the axes.
    #[wasm_bindgen(js_name = "setMax", skip_jsdoc)]
    pub fn set_max(&self, max: f64) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_max(max),
            ChartAxisAccessor::YAxis => lock.y_axis().set_max(max),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_max(max),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_max(max),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the minimum value for an axis.
    ///
    /// Set the minimum bound to be displayed for an axis.
    ///
    /// See {@link ChartAxis#setMax} above for a full explanation and example.
    ///
    /// # Parameters
    ///
    /// `min` - The minimum bound for the axes.
    #[wasm_bindgen(js_name = "setMin", skip_jsdoc)]
    pub fn set_min(&self, min: f64) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_min(min),
            ChartAxisAccessor::YAxis => lock.y_axis().set_min(min),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_min(min),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_min(min),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the increment of the major units in the axis range.
    ///
    /// Note, Excel only supports major/minor units for "Value" axes. In general
    /// you cannot set major/minor units for a X/Category axis even if the
    /// category values are numbers. See [Chart Value and Category Axes] for an
    /// explanation of the difference between Value and Category axes in Excel.
    ///
    /// [Chart Value and Category Axes]:
    ///     crate::chart#chart-value-and-category-axes
    ///
    /// # Parameters
    ///
    /// `value` - The major unit for the axes.
    #[wasm_bindgen(js_name = "setMajorUnit", skip_jsdoc)]
    pub fn set_major_unit(&self, value: f64) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_major_unit(value),
            ChartAxisAccessor::YAxis => lock.y_axis().set_major_unit(value),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_major_unit(value),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_major_unit(value),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the increment of the minor units in the axis range.
    ///
    /// See {@link ChartAxis#setMajorUnit} above for a full explanation and
    /// example.
    ///
    /// # Parameters
    ///
    /// `value` - The major unit for the axes.
    #[wasm_bindgen(js_name = "setMinorUnit", skip_jsdoc)]
    pub fn set_minor_unit(&self, value: f64) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_minor_unit(value),
            ChartAxisAccessor::YAxis => lock.y_axis().set_minor_unit(value),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_minor_unit(value),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_minor_unit(value),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the display unit type such as Thousands, Millions, or other units.
    ///
    /// If the Value axis in your chart has very large numbers you can set the
    /// unit type to one of the following Excel values:
    ///
    /// - Hundreds
    /// - Thousands
    /// - Ten Thousands
    /// - Hundred Thousands
    /// - Millions
    /// - Ten Millions
    /// - Hundred Millions
    /// - Billions
    /// - Trillions
    ///
    /// # Parameters
    ///
    /// - `unit`: A {@link ChartAxisDateUnitType} enum value.
    #[wasm_bindgen(js_name = "setDisplayUnitType", skip_jsdoc)]
    pub fn set_display_unit_type(&self, unit_type: ChartAxisDisplayUnitType) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock
                .x_axis()
                .set_display_unit_type(xlsx::ChartAxisDisplayUnitType::from(unit_type)),
            ChartAxisAccessor::YAxis => lock
                .y_axis()
                .set_display_unit_type(xlsx::ChartAxisDisplayUnitType::from(unit_type)),
            ChartAxisAccessor::X2Axis => lock
                .x2_axis()
                .set_display_unit_type(xlsx::ChartAxisDisplayUnitType::from(unit_type)),
            ChartAxisAccessor::Y2Axis => lock
                .y2_axis()
                .set_display_unit_type(xlsx::ChartAxisDisplayUnitType::from(unit_type)),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Make the display units visible (if they have been set).
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off.
    #[wasm_bindgen(js_name = "setDisplayUnitsVisible", skip_jsdoc)]
    pub fn set_display_units_visible(&self, enable: bool) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_display_units_visible(enable),
            ChartAxisAccessor::YAxis => lock.y_axis().set_display_units_visible(enable),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_display_units_visible(enable),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_display_units_visible(enable),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the major unit type as days, months or years.
    ///
    /// # Parameters
    ///
    /// - `unit`: A {@link ChartAxisDateUnitType} enum value.
    #[wasm_bindgen(js_name = "setMajorUnitDateType", skip_jsdoc)]
    pub fn set_major_unit_date_type(&self, unit_type: ChartAxisDateUnitType) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock
                .x_axis()
                .set_major_unit_date_type(xlsx::ChartAxisDateUnitType::from(unit_type)),
            ChartAxisAccessor::YAxis => lock
                .y_axis()
                .set_major_unit_date_type(xlsx::ChartAxisDateUnitType::from(unit_type)),
            ChartAxisAccessor::X2Axis => lock
                .x2_axis()
                .set_major_unit_date_type(xlsx::ChartAxisDateUnitType::from(unit_type)),
            ChartAxisAccessor::Y2Axis => lock
                .y2_axis()
                .set_major_unit_date_type(xlsx::ChartAxisDateUnitType::from(unit_type)),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the minor unit type as days, months or years.
    ///
    /// # Parameters
    ///
    /// - `unit`: A {@link ChartAxisDateUnitType} enum value.
    #[wasm_bindgen(js_name = "setMinorUnitDateType", skip_jsdoc)]
    pub fn set_minor_unit_date_type(&self, unit_type: ChartAxisDateUnitType) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock
                .x_axis()
                .set_minor_unit_date_type(xlsx::ChartAxisDateUnitType::from(unit_type)),
            ChartAxisAccessor::YAxis => lock
                .y_axis()
                .set_minor_unit_date_type(xlsx::ChartAxisDateUnitType::from(unit_type)),
            ChartAxisAccessor::X2Axis => lock
                .x2_axis()
                .set_minor_unit_date_type(xlsx::ChartAxisDateUnitType::from(unit_type)),
            ChartAxisAccessor::Y2Axis => lock
                .y2_axis()
                .set_minor_unit_date_type(xlsx::ChartAxisDateUnitType::from(unit_type)),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the alignment of the axis labels relative to the tick mark.
    ///
    /// # Parameters
    ///
    /// - `unit`: A {@link ChartAxisDateUnitType} enum value.
    #[wasm_bindgen(js_name = "setLabelAlignment", skip_jsdoc)]
    pub fn set_label_alignment(&self, alignment: ChartAxisLabelAlignment) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock
                .x_axis()
                .set_label_alignment(xlsx::ChartAxisLabelAlignment::from(alignment)),
            ChartAxisAccessor::YAxis => lock
                .y_axis()
                .set_label_alignment(xlsx::ChartAxisLabelAlignment::from(alignment)),
            ChartAxisAccessor::X2Axis => lock
                .x2_axis()
                .set_label_alignment(xlsx::ChartAxisLabelAlignment::from(alignment)),
            ChartAxisAccessor::Y2Axis => lock
                .y2_axis()
                .set_label_alignment(xlsx::ChartAxisLabelAlignment::from(alignment)),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Turn on/off major gridlines for a chart axis.
    ///
    /// Major gridlines are on by default for Y/Value axes but off for
    /// X/Category axes.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default for X axes.
    #[wasm_bindgen(js_name = "setMajorGridlines", skip_jsdoc)]
    pub fn set_major_gridlines(&self, enable: bool) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_major_gridlines(enable),
            ChartAxisAccessor::YAxis => lock.y_axis().set_major_gridlines(enable),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_major_gridlines(enable),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_major_gridlines(enable),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Turn on/off minor gridlines for a chart axis.
    ///
    /// Minor gridlines are off by default.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    #[wasm_bindgen(js_name = "setMinorGridlines", skip_jsdoc)]
    pub fn set_minor_gridlines(&self, enable: bool) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_minor_gridlines(enable),
            ChartAxisAccessor::YAxis => lock.y_axis().set_minor_gridlines(enable),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_minor_gridlines(enable),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_minor_gridlines(enable),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the line formatting for a chart axis major gridlines.
    ///
    /// See the {@link ChartLine} struct for details on the line properties that can
    /// be set.
    ///
    /// # Parameters
    ///
    /// - `line`: A {@link ChartLine} struct reference.
    #[wasm_bindgen(js_name = "setMajorGridlinesLine", skip_jsdoc)]
    pub fn set_major_gridlines_line(&self, line: ChartLine) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_major_gridlines_line(&line.inner),
            ChartAxisAccessor::YAxis => lock.y_axis().set_major_gridlines_line(&line.inner),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_major_gridlines_line(&line.inner),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_major_gridlines_line(&line.inner),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the line formatting for a chart axis minor gridlines.
    ///
    /// See the {@link ChartLine} struct for details on the line properties that can
    /// be set.
    ///
    /// # Parameters
    ///
    /// - `line`: A {@link ChartLine} struct reference.
    #[wasm_bindgen(js_name = "setMinorGridlinesLine", skip_jsdoc)]
    pub fn set_minor_gridlines_line(&self, line: ChartLine) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_minor_gridlines_line(&line.inner),
            ChartAxisAccessor::YAxis => lock.y_axis().set_minor_gridlines_line(&line.inner),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_minor_gridlines_line(&line.inner),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_minor_gridlines_line(&line.inner),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the label position for the axis.
    ///
    /// The label position defines where the values/categories for the axis are
    /// displayed. The position is controlled via {@link ChartAxisLabelPosition} enum.
    ///
    /// # Parameters
    ///
    /// - `position`: A {@link ChartAxisLabelPosition} enum value.
    #[wasm_bindgen(js_name = "setLabelPosition", skip_jsdoc)]
    pub fn set_label_position(&self, position: ChartAxisLabelPosition) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock
                .x_axis()
                .set_label_position(xlsx::ChartAxisLabelPosition::from(position)),
            ChartAxisAccessor::YAxis => lock
                .y_axis()
                .set_label_position(xlsx::ChartAxisLabelPosition::from(position)),
            ChartAxisAccessor::X2Axis => lock
                .x2_axis()
                .set_label_position(xlsx::ChartAxisLabelPosition::from(position)),
            ChartAxisAccessor::Y2Axis => lock
                .y2_axis()
                .set_label_position(xlsx::ChartAxisLabelPosition::from(position)),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the axis position on or between the tick marks.
    ///
    /// In Excel there are two "Axis position" options for Category axes: "On
    /// tick marks" and "Between tick marks". This property has different
    /// default value for different chart types and isn't available for some
    /// chart types like Scatter. The `set_position_between_ticks()` method can
    /// be used to change the default value.
    ///
    /// Note, this property is only applicable to Category axes, see [Chart
    /// Value and Category Axes] for an explanation of the difference between
    /// Value and Category axes in Excel.
    ///
    /// [Chart Value and Category Axes]:
    ///     crate::chart#chart-value-and-category-axes
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. Its default value depends on the
    ///   chart type.
    #[wasm_bindgen(js_name = "setPositionBetweenTicks", skip_jsdoc)]
    pub fn set_position_between_ticks(&self, enable: bool) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_position_between_ticks(enable),
            ChartAxisAccessor::YAxis => lock.y_axis().set_position_between_ticks(enable),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_position_between_ticks(enable),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_position_between_ticks(enable),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the interval of the axis labels.
    ///
    /// Set the interval of the axis labels for Category axes. This value is 1
    /// by default, i.e., there is one label shown per category. If needed it
    /// can be set to another value.
    ///
    /// Note, this property is only applicable to Category axes, see [Chart
    /// Value and Category Axes] for an explanation of the difference between
    /// Value and Category axes in Excel.
    ///
    /// # Parameters
    ///
    /// - `interval`: The interval for the category labels. The default is 1.
    #[wasm_bindgen(js_name = "setLabelInterval", skip_jsdoc)]
    pub fn set_label_interval(&self, interval: u16) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_label_interval(interval),
            ChartAxisAccessor::YAxis => lock.y_axis().set_label_interval(interval),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_label_interval(interval),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_label_interval(interval),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the interval of the axis ticks.
    ///
    /// Set the interval of the axis ticks for Category axes. This value is 1
    /// by default, i.e., there is one tick shown per category. If needed it
    /// can be set to another value.
    ///
    /// Note, this property is only applicable to Category axes, see [Chart
    /// Value and Category Axes] for an explanation of the difference between
    /// Value and Category axes in Excel.
    ///
    /// # Parameters
    ///
    /// - `interval`: The interval for the category ticks. The default is 1.
    #[wasm_bindgen(js_name = "setTickInterval", skip_jsdoc)]
    pub fn set_tick_interval(&self, interval: u16) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_tick_interval(interval),
            ChartAxisAccessor::YAxis => lock.y_axis().set_tick_interval(interval),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_tick_interval(interval),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_tick_interval(interval),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the type of major tick for the axis.
    ///
    /// Excel supports 4 types of tick position:
    ///
    /// - None
    /// - Inside only
    /// - Outside only
    /// - Cross - inside and outside
    ///
    /// # Parameters
    ///
    /// - `tick_type`: a {@link ChartAxisTickType} enum value.
    #[wasm_bindgen(js_name = "setMajorTickType", skip_jsdoc)]
    pub fn set_major_tick_type(&self, tick_type: ChartAxisTickType) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock
                .x_axis()
                .set_major_tick_type(xlsx::ChartAxisTickType::from(tick_type)),
            ChartAxisAccessor::YAxis => lock
                .y_axis()
                .set_major_tick_type(xlsx::ChartAxisTickType::from(tick_type)),
            ChartAxisAccessor::X2Axis => lock
                .x2_axis()
                .set_major_tick_type(xlsx::ChartAxisTickType::from(tick_type)),
            ChartAxisAccessor::Y2Axis => lock
                .y2_axis()
                .set_major_tick_type(xlsx::ChartAxisTickType::from(tick_type)),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the type of minor tick for the axis.
    ///
    /// See {@link ChartAxis#setMajorTickType} above for an explanation and
    /// example.
    ///
    /// # Parameters
    ///
    /// - `tick_type`: a {@link ChartAxisTickType} enum value.
    #[wasm_bindgen(js_name = "setMinorTickType", skip_jsdoc)]
    pub fn set_minor_tick_type(&self, tick_type: ChartAxisTickType) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock
                .x_axis()
                .set_minor_tick_type(xlsx::ChartAxisTickType::from(tick_type)),
            ChartAxisAccessor::YAxis => lock
                .y_axis()
                .set_minor_tick_type(xlsx::ChartAxisTickType::from(tick_type)),
            ChartAxisAccessor::X2Axis => lock
                .x2_axis()
                .set_minor_tick_type(xlsx::ChartAxisTickType::from(tick_type)),
            ChartAxisAccessor::Y2Axis => lock
                .y2_axis()
                .set_minor_tick_type(xlsx::ChartAxisTickType::from(tick_type)),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the log base of the axis range.
    ///
    /// This property is only applicable to value axes, see [Chart Value and
    /// Category Axes] for an explanation of the difference between Value and
    /// Category axes in Excel.
    ///
    /// [Chart Value and Category Axes]:
    ///     crate::chart#chart-value-and-category-axes
    ///
    /// # Parameters
    ///
    /// - `base`: The logarithm base. Should be >= 2.
    #[wasm_bindgen(js_name = "setLogBase", skip_jsdoc)]
    pub fn set_log_base(&self, base: u16) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_log_base(base),
            ChartAxisAccessor::YAxis => lock.y_axis().set_log_base(base),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_log_base(base),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_log_base(base),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Hide the chart axis.
    ///
    /// Hide the number or label section of the chart axis.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off.
    #[wasm_bindgen(js_name = "setHidden", skip_jsdoc)]
    pub fn set_hidden(&self, enable: bool) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_hidden(enable),
            ChartAxisAccessor::YAxis => lock.y_axis().set_hidden(enable),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_hidden(enable),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_hidden(enable),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
    /// Set the manual position of the chart axis label.
    ///
    /// This method is used to simulate manual positioning of a chart axis
    /// label. See {@link ChartLayout} for more details.
    ///
    /// # Parameters
    ///
    /// - `layout`: A {@link ChartLayout} struct reference.
    #[wasm_bindgen(js_name = "setLabelLayout", skip_jsdoc)]
    pub fn set_label_layout(&self, layout: ChartLayout) -> ChartAxis {
        let mut lock = self.parent.lock().unwrap();
        match self.accessor {
            ChartAxisAccessor::XAxis => lock.x_axis().set_label_layout(&layout.inner),
            ChartAxisAccessor::YAxis => lock.y_axis().set_label_layout(&layout.inner),
            ChartAxisAccessor::X2Axis => lock.x2_axis().set_label_layout(&layout.inner),
            ChartAxisAccessor::Y2Axis => lock.y2_axis().set_label_layout(&layout.inner),
        }
        ChartAxis {
            parent: Arc::clone(&self.parent),
        }
    }
}
