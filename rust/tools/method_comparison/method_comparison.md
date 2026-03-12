# Method Comparison Report: rustxlsxwriter to wasm-xlsxwriter Migration

## Summary
  - ✅ Fully Migrated Structs: 46
  - ⚠️ Partially Migrated Structs: 14
  - ❌ Not Migrated Structs: 2
  - ✅ Migrated Enums: 47
  - ❌ Not Migrated Enums: 5
  - ✅ Total Migrated Methods: 512
  - ❌ Total Not Migrated Methods: 125
  - ✅ Total Migrated Functions: 69
  - ❌ Total Not Migrated Functions: 8
## Details of Structs
  ### ✅ Button
  ### ✅ ChartAxis
  ### ✅ ChartDataTable
  ### ✅ ChartFont
  ### ✅ ChartFormat
  ### ✅ ChartGradientFill
  ### ✅ ChartGradientStop
  ### ✅ ChartLayout
  ### ✅ ChartLegend
  ### ✅ ChartLine
  ### ✅ ChartMarker
  ### ✅ ChartPatternFill
  ### ✅ ChartPoint
  ### ✅ ChartRange
  ### ✅ ChartSolidFill
  ### ✅ ChartTitle
  ### ✅ ChartTrendline
  ### ✅ ConditionalFormat2ColorScale
  ### ✅ ConditionalFormat3ColorScale
  ### ✅ ConditionalFormatAverage
  ### ✅ ConditionalFormatBlank
  ### ✅ ConditionalFormatCustomIcon
  ### ✅ ConditionalFormatDataBar
  ### ✅ ConditionalFormatDate
  ### ✅ ConditionalFormatDuplicate
  ### ✅ ConditionalFormatError
  ### ✅ ConditionalFormatFormula
  ### ✅ ConditionalFormatIconSet
  ### ✅ ConditionalFormatText
  ### ✅ ConditionalFormatTop
  ### ✅ ExcelDateTime
  ### ✅ Format
  ### ✅ Formula
  ### ✅ Note
  ### ✅ ProtectionOptions
  ### ✅ ShapeFont
  ### ✅ ShapeFormat
  ### ✅ ShapeGradientFill
  ### ✅ ShapeGradientStop
  ### ✅ ShapeLine
  ### ✅ ShapePatternFill
  ### ✅ ShapeSolidFill
  ### ✅ ShapeText
  ### ✅ Sparkline
  ### ✅ TableColumn
  ### ✅ Url
  ### ⚠️ Chart
    Summary
      - Migrated methods: 30
      - Not migrated methods: 4
      - Migrated functions: 10
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - add_series
      - chart_area
      - plot_area
      - validate
  ### ⚠️ ChartArea
    Summary
      - Migrated methods: 1
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Functions Not Yet Migrated
      - new
  ### ⚠️ ChartDataLabel
    Summary
      - Migrated methods: 15
      - Not migrated methods: 1
      - Migrated functions: 1
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - set_separator
  ### ⚠️ ChartErrorBars
    Summary
      - Migrated methods: 3
      - Not migrated methods: 1
      - Migrated functions: 1
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - set_type
  ### ⚠️ ChartPlotArea
    Summary
      - Migrated methods: 2
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Functions Not Yet Migrated
      - new
  ### ⚠️ ChartSeries
    Summary
      - Migrated methods: 18
      - Not migrated methods: 1
      - Migrated functions: 1
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - set_point_colors
  ### ⚠️ ConditionalFormatCell
    Summary
      - Migrated methods: 3
      - Not migrated methods: 1
      - Migrated functions: 1
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - set_rule
  ### ⚠️ DataValidation
    Summary
      - Migrated methods: 9
      - Not migrated methods: 15
      - Migrated functions: 1
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - allow_date
      - allow_date_formula
      - allow_decimal_number
      - allow_decimal_number_formula
      - allow_list_strings
      - allow_text_length
      - allow_text_length_formula
      - allow_time
      - allow_time_formula
      - allow_whole_number
      - allow_whole_number_formula
      - set_error_message
      - set_error_title
      - set_input_message
      - set_input_title
  ### ⚠️ DocProperties
    Summary
      - Migrated methods: 10
      - Not migrated methods: 2
      - Migrated functions: 1
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - set_creation_datetime
      - set_custom_property
  ### ⚠️ FilterCondition
    Summary
      - Migrated methods: 2
      - Not migrated methods: 2
      - Migrated functions: 1
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - add_custom_filter
      - add_list_filter
  ### ⚠️ Image
    Summary
      - Migrated methods: 8
      - Not migrated methods: 5
      - Migrated functions: 1
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - height
      - height_dpi
      - set_url
      - width
      - width_dpi
    ❌ Functions Not Yet Migrated
      - new_from_buffer
  ### ⚠️ Table
    Summary
      - Migrated methods: 8
      - Not migrated methods: 4
      - Migrated functions: 1
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - set_alt_text
      - set_alt_text_title
      - set_autofilter
      - set_last_column
  ### ⚠️ Workbook
    Summary
      - Migrated methods: 7
      - Not migrated methods: 13
      - Migrated functions: 1
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - add_chartsheet
      - add_vba_project
      - add_vba_project_with_signature
      - push_worksheet
      - save
      - save_to_buffer
      - save_to_writer
      - set_vba_name
      - use_custom_theme
      - use_excel_2023_theme
      - use_zip_large_file
      - worksheets
      - worksheets_mut
  ### ⚠️ Worksheet
    Summary
      - Migrated methods: 80
      - Not migrated methods: 66
      - Migrated functions: 0
      - Not migrated functions: 2
    ❌ Methods Not Yet Migrated
      - add_conditional_format
      - add_data_validation
      - add_sparkline
      - add_sparkline_group
      - filter_automatic_off
      - filter_column
      - group_columns
      - group_columns_collapsed
      - group_rows_collapsed
      - group_symbols_above
      - group_symbols_to_left
      - hide_unused_rows
      - ignore_error
      - ignore_error_range
      - insert_background_image
      - insert_button
      - insert_button_with_offset
      - insert_checkbox
      - insert_checkbox_with_format
      - insert_shape
      - insert_shape_with_offset
      - protect_with_options
      - protect_with_password
      - set_autofit_max_row
      - set_autofit_max_width
      - set_cell_format
      - set_column_autofit_width
      - set_column_format
      - set_column_hidden
      - set_column_range_format
      - set_column_range_hidden
      - set_column_range_width_pixels
      - set_default_format
      - set_default_note_author
      - set_default_row_height
      - set_default_row_height_pixels
      - set_first_tab
      - set_formula_result
      - set_formula_result_default
      - set_header_footer_align_with_page
      - set_header_footer_scale_with_doc
      - set_infinity_value
      - set_nan_value
      - set_neg_infinity_value
      - set_page_breaks
      - set_page_order
      - set_right_to_left
      - set_row_format
      - set_row_hidden
      - set_row_unhidden
      - set_selected
      - set_selection
      - set_tab_color
      - set_top_left_cell
      - set_vba_name
      - set_vertical_page_breaks
      - set_very_hidden
      - set_view_normal
      - set_view_page_break_preview
      - set_view_page_layout
      - set_zoom
      - set_zoom_to_fit
      - show_all_notes
      - unprotect_range
      - unprotect_range_with_options
      - write_dynamic_formula
    ❌ Functions Not Yet Migrated
      - new
      - new_chartsheet
  ### ❌ FilterData
    Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 2
    ❌ Functions Not Yet Migrated
      - new_number_and_criteria
      - new_string_and_criteria
  ### ❌ Shape
    Summary
      - Migrated methods: 0
      - Not migrated methods: 10
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_alt_text
      - set_font
      - set_format
      - set_height
      - set_object_movement
      - set_text
      - set_text_link
      - set_text_options
      - set_url
      - set_width
    ❌ Functions Not Yet Migrated
      - textbox

## ✅ Migrated Enums
  - ChartAxisCrossing
  - ChartAxisDateUnitType
  - ChartAxisDisplayUnitType
  - ChartAxisLabelAlignment
  - ChartAxisLabelPosition
  - ChartAxisTickType
  - ChartDataLabelPosition
  - ChartEmptyCells
  - ChartErrorBarsDirection
  - ChartGradientFillType
  - ChartLegendPosition
  - ChartLineDashType
  - ChartMarkerType
  - ChartPatternFillType
  - ChartTrendlineType
  - ChartType
  - Color
  - ConditionalFormatAverageRule
  - ConditionalFormatDataBarAxisPosition
  - ConditionalFormatDataBarDirection
  - ConditionalFormatDateRule
  - ConditionalFormatIconType
  - ConditionalFormatTextRule
  - ConditionalFormatTopRule
  - ConditionalFormatType
  - DataValidationErrorStyle
  - FilterCriteria
  - FontScheme
  - FormatAlign
  - FormatBorder
  - FormatDiagonalBorder
  - FormatPattern
  - FormatScript
  - FormatUnderline
  - HeaderImagePosition
  - IgnoreError
  - ObjectMovement
  - ShapeGradientFillType
  - ShapeLineDashType
  - ShapePatternFillType
  - ShapeTextDirection
  - ShapeTextHorizontalAlignment
  - ShapeTextVerticalAlignment
  - SparklineType
  - TableFunction
  - TableStyle
  - XlsxError

## ❌ Not Migrated Enums
  - ChartAxisAccessor
  - ChartErrorBarsType
  - ConditionalFormatCellRule
  - DataValidationRule
  - ExcelData
