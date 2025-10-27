# Method Comparison Report: rustxlsxwriter to wasm-xlsxwriter Migration

## Summary
  - ✅ Fully Migrated Structs: 18
  - ⚠️ Partially Migrated Structs: 10
  - ❌ Not Migrated Structs: 34
  - ✅ Migrated Enums: 19
  - ❌ Not Migrated Enums: 32
  - ✅ Total Migrated Methods: 262
  - ❌ Total Not Migrated Methods: 369
  - ✅ Total Migrated Functions: 39
  - ❌ Total Not Migrated Functions: 38
## Details of Structs
  ### ✅ ChartDataLabel
  ### ✅ ChartFont
  ### ✅ ChartFormat
  ### ✅ ChartGradientFill
  ### ✅ ChartGradientStop
  ### ✅ ChartLayout
  ### ✅ ChartLine
  ### ✅ ChartMarker
  ### ✅ ChartPatternFill
  ### ✅ ChartPoint
  ### ✅ ChartRange
  ### ✅ ChartSolidFill
  ### ✅ ChartTitle
  ### ✅ ExcelDateTime
  ### ✅ Formula
  ### ✅ Note
  ### ✅ TableColumn
  ### ✅ Url
  ### ⚠️ Chart
    Summary
      - Migrated methods: 11
      - Not migrated methods: 23
      - Migrated functions: 10
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - add_series
      - chart_area
      - combine
      - plot_area
      - set_data_table
      - set_decorative
      - set_down_bar_format
      - set_drop_lines
      - set_drop_lines_format
      - set_high_low_lines
      - set_high_low_lines_format
      - set_hole_size
      - set_object_movement
      - set_rotation
      - set_scale_height
      - set_scale_width
      - set_style
      - set_up_bar_format
      - set_up_down_bars
      - show_empty_cells_as
      - show_hidden_data
      - show_na_as_empty_cell
      - validate
  ### ⚠️ ChartAxis
    Summary
      - Migrated methods: 6
      - Not migrated methods: 29
      - Migrated functions: 0
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - set_automatic_axis
      - set_crossing
      - set_date_axis
      - set_display_unit_type
      - set_display_units_visible
      - set_format
      - set_hidden
      - set_label_alignment
      - set_label_interval
      - set_label_layout
      - set_label_position
      - set_log_base
      - set_major_gridlines
      - set_major_gridlines_line
      - set_major_tick_type
      - set_major_unit
      - set_major_unit_date_type
      - set_max_date
      - set_min_date
      - set_minor_gridlines
      - set_minor_gridlines_line
      - set_minor_tick_type
      - set_minor_unit
      - set_minor_unit_date_type
      - set_name_format
      - set_position_between_ticks
      - set_reverse
      - set_text_axis
      - set_tick_interval
  ### ⚠️ ChartLegend
    Summary
      - Migrated methods: 5
      - Not migrated methods: 2
      - Migrated functions: 0
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - delete_entries
      - set_layout
  ### ⚠️ ChartSeries
    Summary
      - Migrated methods: 7
      - Not migrated methods: 12
      - Migrated functions: 1
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - delete_from_legend
      - set_custom_data_labels
      - set_gap
      - set_invert_if_negative
      - set_invert_if_negative_color
      - set_overlap
      - set_point_colors
      - set_secondary_axis
      - set_smooth
      - set_trendline
      - set_x_error_bars
      - set_y_error_bars
  ### ⚠️ DocProperties
    Summary
      - Migrated methods: 10
      - Not migrated methods: 2
      - Migrated functions: 1
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - set_creation_datetime
      - set_custom_property
  ### ⚠️ Format
    Summary
      - Migrated methods: 37
      - Not migrated methods: 14
      - Migrated functions: 1
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - merge
      - set_checkbox
      - set_num_format_index
      - set_reading_direction
      - set_shrink
      - unset_bold
      - unset_checkbox
      - unset_font_strikethrough
      - unset_hidden
      - unset_hyperlink_style
      - unset_italic
      - unset_quote_prefix
      - unset_shrink
      - unset_text_wrap
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
      - Migrated methods: 6
      - Not migrated methods: 11
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
      - use_zip_large_file
      - worksheets
      - worksheets_mut
  ### ⚠️ Worksheet
    Summary
      - Migrated methods: 78
      - Not migrated methods: 65
      - Migrated functions: 0
      - Not migrated functions: 2
    ❌ Methods Not Yet Migrated
      - add_conditional_format
      - add_data_validation
      - add_sparkline
      - add_sparkline_group
      - autofit_to_max_width
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
      - set_cell_format
      - set_column_autofit_width
      - set_column_format
      - set_column_hidden
      - set_column_range_format
      - set_column_range_hidden
      - set_column_range_width_pixels
      - set_default_note_author
      - set_default_row_height
      - set_default_row_height_pixels
      - set_first_tab
      - set_formula_result
      - set_formula_result_default
      - set_header_footer_align_with_page
      - set_header_footer_scale_with_doc
      - set_infinity_value
      - set_margins
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
  ### ❌ Button
    Summary
      - Migrated methods: 0
      - Not migrated methods: 6
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_alt_text
      - set_caption
      - set_height
      - set_macro
      - set_object_movement
      - set_width
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ChartArea
    Summary
      - Migrated methods: 0
      - Not migrated methods: 1
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_format
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ChartDataTable
    Summary
      - Migrated methods: 0
      - Not migrated methods: 6
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_font
      - set_format
      - show_horizontal_borders
      - show_legend_keys
      - show_outline_borders
      - show_vertical_borders
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ChartErrorBars
    Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_direction
      - set_end_cap
      - set_format
      - set_type
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ChartPlotArea
    Summary
      - Migrated methods: 0
      - Not migrated methods: 2
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_format
      - set_layout
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ChartTrendline
    Summary
      - Migrated methods: 0
      - Not migrated methods: 11
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - delete_from_legend
      - display_equation
      - display_r_squared
      - set_backward_period
      - set_format
      - set_forward_period
      - set_intercept
      - set_label_font
      - set_label_format
      - set_name
      - set_type
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ConditionalFormat2ColorScale
    Summary
      - Migrated methods: 0
      - Not migrated methods: 6
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_maximum
      - set_maximum_color
      - set_minimum
      - set_minimum_color
      - set_multi_range
      - set_stop_if_true
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ConditionalFormat3ColorScale
    Summary
      - Migrated methods: 0
      - Not migrated methods: 8
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_maximum
      - set_maximum_color
      - set_midpoint
      - set_midpoint_color
      - set_minimum
      - set_minimum_color
      - set_multi_range
      - set_stop_if_true
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ConditionalFormatAverage
    Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_format
      - set_multi_range
      - set_rule
      - set_stop_if_true
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ConditionalFormatBlank
    Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - invert
      - set_format
      - set_multi_range
      - set_stop_if_true
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ConditionalFormatCell
    Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_format
      - set_multi_range
      - set_rule
      - set_stop_if_true
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ConditionalFormatCustomIcon
    Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_greater_than
      - set_icon_type
      - set_no_icon
      - set_rule
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ConditionalFormatDataBar
    Summary
      - Migrated methods: 0
      - Not migrated methods: 14
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_axis_color
      - set_axis_position
      - set_bar_only
      - set_border_color
      - set_border_off
      - set_direction
      - set_fill_color
      - set_maximum
      - set_minimum
      - set_multi_range
      - set_negative_border_color
      - set_negative_fill_color
      - set_solid_fill
      - set_stop_if_true
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ConditionalFormatDate
    Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_format
      - set_multi_range
      - set_rule
      - set_stop_if_true
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ConditionalFormatDuplicate
    Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - invert
      - set_format
      - set_multi_range
      - set_stop_if_true
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ConditionalFormatError
    Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - invert
      - set_format
      - set_multi_range
      - set_stop_if_true
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ConditionalFormatFormula
    Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_format
      - set_multi_range
      - set_rule
      - set_stop_if_true
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ConditionalFormatIconSet
    Summary
      - Migrated methods: 0
      - Not migrated methods: 6
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - reverse_icons
      - set_icon_type
      - set_icons
      - set_multi_range
      - set_stop_if_true
      - show_icons_only
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ConditionalFormatText
    Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_format
      - set_multi_range
      - set_rule
      - set_stop_if_true
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ConditionalFormatTop
    Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_format
      - set_multi_range
      - set_rule
      - set_stop_if_true
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ DataValidation
    Summary
      - Migrated methods: 0
      - Not migrated methods: 24
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - allow_any_value
      - allow_custom
      - allow_date
      - allow_date_formula
      - allow_decimal_number
      - allow_decimal_number_formula
      - allow_list_formula
      - allow_list_strings
      - allow_text_length
      - allow_text_length_formula
      - allow_time
      - allow_time_formula
      - allow_whole_number
      - allow_whole_number_formula
      - ignore_blank
      - set_error_message
      - set_error_style
      - set_error_title
      - set_input_message
      - set_input_title
      - set_multi_range
      - show_dropdown
      - show_error_message
      - show_input_message
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ FilterCondition
    Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - add_custom_boolean_or
      - add_custom_filter
      - add_list_blanks_filter
      - add_list_filter
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ FilterData
    Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 2
    ❌ Functions Not Yet Migrated
      - new_number_and_criteria
      - new_string_and_criteria
  ### ❌ ProtectionOptions
    Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Functions Not Yet Migrated
      - new
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
  ### ❌ ShapeFont
    Summary
      - Migrated methods: 0
      - Not migrated methods: 11
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_bold
      - set_character_set
      - set_color
      - set_italic
      - set_name
      - set_pitch_family
      - set_right_to_left
      - set_size
      - set_strikethrough
      - set_underline
      - unset_bold
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ShapeFormat
    Summary
      - Migrated methods: 0
      - Not migrated methods: 6
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_gradient_fill
      - set_line
      - set_no_fill
      - set_no_line
      - set_pattern_fill
      - set_solid_fill
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ShapeGradientFill
    Summary
      - Migrated methods: 0
      - Not migrated methods: 3
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_angle
      - set_gradient_stops
      - set_type
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ShapeGradientStop
    Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ShapeLine
    Summary
      - Migrated methods: 0
      - Not migrated methods: 5
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_color
      - set_dash_type
      - set_hidden
      - set_transparency
      - set_width
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ShapePatternFill
    Summary
      - Migrated methods: 0
      - Not migrated methods: 3
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_background_color
      - set_foreground_color
      - set_pattern
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ShapeSolidFill
    Summary
      - Migrated methods: 0
      - Not migrated methods: 2
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_color
      - set_transparency
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ ShapeText
    Summary
      - Migrated methods: 0
      - Not migrated methods: 3
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_direction
      - set_horizontal_alignment
      - set_vertical_alignment
    ❌ Functions Not Yet Migrated
      - new
  ### ❌ Sparkline
    Summary
      - Migrated methods: 0
      - Not migrated methods: 27
      - Migrated functions: 0
      - Not migrated functions: 1
    ❌ Methods Not Yet Migrated
      - set_column_order
      - set_custom_max
      - set_custom_min
      - set_date_range
      - set_first_point_color
      - set_group_max
      - set_group_min
      - set_high_point_color
      - set_last_point_color
      - set_line_weight
      - set_low_point_color
      - set_markers_color
      - set_negative_points_color
      - set_range
      - set_right_to_left
      - set_sparkline_color
      - set_style
      - set_type
      - show_axis
      - show_empty_cells_as
      - show_first_point
      - show_hidden_data
      - show_high_point
      - show_last_point
      - show_low_point
      - show_markers
      - show_negative_points
    ❌ Functions Not Yet Migrated
      - new

## ✅ Migrated Enums
  - ChartDataLabelPosition
  - ChartGradientFillType
  - ChartLegendPosition
  - ChartLineDashType
  - ChartMarkerType
  - ChartPatternFillType
  - ChartType
  - Color
  - FormatAlign
  - FormatBorder
  - FormatDiagonalBorder
  - FormatPattern
  - FormatScript
  - FormatUnderline
  - HeaderImagePosition
  - ObjectMovement
  - TableFunction
  - TableStyle
  - XlsxError

## ❌ Not Migrated Enums
  - AxisType
  - ChartAxisCrossing
  - ChartAxisDateUnitType
  - ChartAxisDisplayUnitType
  - ChartAxisLabelAlignment
  - ChartAxisLabelPosition
  - ChartAxisTickType
  - ChartEmptyCells
  - ChartErrorBarsDirection
  - ChartErrorBarsType
  - ChartTrendlineType
  - ConditionalFormatAverageRule
  - ConditionalFormatCellRule
  - ConditionalFormatDataBarAxisPosition
  - ConditionalFormatDataBarDirection
  - ConditionalFormatDateRule
  - ConditionalFormatIconType
  - ConditionalFormatTextRule
  - ConditionalFormatTopRule
  - ConditionalFormatType
  - DataValidationErrorStyle
  - DataValidationRule
  - ExcelData
  - FilterCriteria
  - IgnoreError
  - ShapeGradientFillType
  - ShapeLineDashType
  - ShapePatternFillType
  - ShapeTextDirection
  - ShapeTextHorizontalAlignment
  - ShapeTextVerticalAlignment
  - SparklineType
