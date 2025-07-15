# Method Comparison Report: rustxlsxwriter to wasm-xlsxwriter Migration

## ✅ Migrated Structs
  - Chart
  - ChartAxis
  - ChartDataLabel
  - ChartFont
  - ChartFormat
  - ChartLegend
  - ChartLine
  - ChartMarker
  - ChartPoint
  - ChartSeries
  - ChartSolidFill
  - ChartTitle
  - DocProperties
  - ExcelDateTime
  - Formula
  - Image
  - Note
  - Table
  - TableColumn
  - Url

## ⚠️ Partially Migrated Structs
  ### ChartRange
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 1
      - Not migrated functions: 1
    #### ❌ Functions Not Yet Migrated
      - new_from_range
  ### Format
    #### Summary
      - Migrated methods: 36
      - Not migrated methods: 15
      - Migrated functions: 1
      - Not migrated functions: 0
    #### ❌ Methods Not Yet Migrated
      - unset_checkbox
      - set_shrink
      - unset_hyperlink_style
      - set_font_script
      - unset_italic
      - unset_text_wrap
      - unset_shrink
      - unset_font_strikethrough
      - set_checkbox
      - merge
      - unset_bold
      - unset_quote_prefix
      - set_reading_direction
      - set_num_format_index
      - unset_hidden
  ### Workbook
    #### Summary
      - Migrated methods: 6
      - Not migrated methods: 11
      - Migrated functions: 1
      - Not migrated functions: 0
    #### ❌ Methods Not Yet Migrated
      - worksheets_mut
      - set_vba_name
      - worksheets
      - add_chartsheet
      - save_to_buffer
      - add_vba_project
      - push_worksheet
      - save
      - save_to_writer
      - add_vba_project_with_signature
      - use_zip_large_file
  ### Worksheet
    #### Summary
      - Migrated methods: 77
      - Not migrated methods: 64
      - Migrated functions: 0
      - Not migrated functions: 2
    #### ❌ Methods Not Yet Migrated
      - filter_column
      - unprotect_range_with_options
      - set_nan_value
      - set_header_footer_align_with_page
      - show_all_notes
      - set_column_range_format
      - set_selection
      - insert_checkbox_with_format
      - set_infinity_value
      - group_columns
      - set_view_page_break_preview
      - protect_with_password
      - set_cell_format
      - set_default_note_author
      - set_header_footer_scale_with_doc
      - set_default_row_height
      - set_neg_infinity_value
      - group_columns_collapsed
      - set_top_left_cell
      - set_vba_name
      - add_sparkline
      - set_view_page_layout
      - set_column_range_width_pixels
      - set_row_format
      - set_first_tab
      - set_column_hidden
      - autofit_to_max_width
      - write_dynamic_formula
      - set_selected
      - set_right_to_left
      - add_data_validation
      - set_very_hidden
      - set_view_normal
      - set_zoom
      - set_page_breaks
      - set_column_range_hidden
      - add_conditional_format
      - hide_unused_rows
      - set_vertical_page_breaks
      - set_row_hidden
      - add_sparkline_group
      - ignore_error_range
      - insert_shape
      - group_rows_collapsed
      - group_symbols_above
      - set_margins
      - insert_shape_with_offset
      - set_formula_result
      - insert_checkbox
      - unprotect_range
      - insert_button_with_offset
      - set_tab_color
      - insert_button
      - set_page_order
      - set_default_row_height_pixels
      - ignore_error
      - set_formula_result_default
      - group_symbols_to_left
      - set_column_autofit_width
      - filter_automatic_off
      - set_column_format
      - insert_background_image
      - set_row_unhidden
      - protect_with_options
    #### ❌ Functions Not Yet Migrated
      - new_chartsheet
      - new

## ❌ Not Migrated Structs
  ### App
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### Button
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 6
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_height
      - set_width
      - set_caption
      - set_macro
      - set_object_movement
      - set_alt_text
    #### ❌ Functions Not Yet Migrated
      - new
  ### ChartArea
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 1
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_format
    #### ❌ Functions Not Yet Migrated
      - new
  ### ChartDataTable
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 6
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_format
      - set_font
      - show_horizontal_borders
      - show_vertical_borders
      - show_outline_borders
      - show_legend_keys
    #### ❌ Functions Not Yet Migrated
      - new
  ### ChartErrorBars
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_format
      - set_type
      - set_direction
      - set_end_cap
    #### ❌ Functions Not Yet Migrated
      - new
  ### ChartGradientFill
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 3
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_type
      - set_gradient_stops
      - set_angle
    #### ❌ Functions Not Yet Migrated
      - new
  ### ChartGradientStop
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Functions Not Yet Migrated
      - new
  ### ChartLayout
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 2
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_dimensions
      - set_offset
    #### ❌ Functions Not Yet Migrated
      - new
  ### ChartPatternFill
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 3
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_pattern
      - set_foreground_color
      - set_background_color
    #### ❌ Functions Not Yet Migrated
      - new
  ### ChartPlotArea
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 2
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_format
      - set_layout
    #### ❌ Functions Not Yet Migrated
      - new
  ### ChartTrendline
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 11
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_forward_period
      - set_label_font
      - set_type
      - display_equation
      - display_r_squared
      - set_label_format
      - delete_from_legend
      - set_format
      - set_backward_period
      - set_name
      - set_intercept
    #### ❌ Functions Not Yet Migrated
      - new
  ### Color
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### Comment
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### ConditionalFormat2ColorScale
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 6
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_stop_if_true
      - set_maximum
      - set_maximum_color
      - set_minimum
      - set_minimum_color
      - set_multi_range
    #### ❌ Functions Not Yet Migrated
      - new
  ### ConditionalFormat3ColorScale
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 8
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_maximum_color
      - set_minimum
      - set_maximum
      - set_midpoint_color
      - set_multi_range
      - set_stop_if_true
      - set_minimum_color
      - set_midpoint
    #### ❌ Functions Not Yet Migrated
      - new
  ### ConditionalFormatAverage
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_format
      - set_rule
      - set_stop_if_true
      - set_multi_range
    #### ❌ Functions Not Yet Migrated
      - new
  ### ConditionalFormatBlank
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - invert
      - set_format
      - set_multi_range
      - set_stop_if_true
    #### ❌ Functions Not Yet Migrated
      - new
  ### ConditionalFormatCell
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_rule
      - set_multi_range
      - set_stop_if_true
      - set_format
    #### ❌ Functions Not Yet Migrated
      - new
  ### ConditionalFormatCustomIcon
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_rule
      - set_greater_than
      - set_icon_type
      - set_no_icon
    #### ❌ Functions Not Yet Migrated
      - new
  ### ConditionalFormatDataBar
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 14
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_maximum
      - set_negative_fill_color
      - set_solid_fill
      - set_negative_border_color
      - set_stop_if_true
      - set_border_off
      - set_axis_position
      - set_border_color
      - set_direction
      - set_multi_range
      - set_fill_color
      - set_axis_color
      - set_bar_only
      - set_minimum
    #### ❌ Functions Not Yet Migrated
      - new
  ### ConditionalFormatDate
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_stop_if_true
      - set_multi_range
      - set_rule
      - set_format
    #### ❌ Functions Not Yet Migrated
      - new
  ### ConditionalFormatDuplicate
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_format
      - set_multi_range
      - invert
      - set_stop_if_true
    #### ❌ Functions Not Yet Migrated
      - new
  ### ConditionalFormatError
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_format
      - invert
      - set_multi_range
      - set_stop_if_true
    #### ❌ Functions Not Yet Migrated
      - new
  ### ConditionalFormatFormula
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_format
      - set_stop_if_true
      - set_multi_range
      - set_rule
    #### ❌ Functions Not Yet Migrated
      - new
  ### ConditionalFormatIconSet
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 6
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_stop_if_true
      - set_multi_range
      - show_icons_only
      - set_icon_type
      - reverse_icons
      - set_icons
    #### ❌ Functions Not Yet Migrated
      - new
  ### ConditionalFormatText
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_multi_range
      - set_stop_if_true
      - set_rule
      - set_format
    #### ❌ Functions Not Yet Migrated
      - new
  ### ConditionalFormatTop
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_multi_range
      - set_rule
      - set_format
      - set_stop_if_true
    #### ❌ Functions Not Yet Migrated
      - new
  ### ConditionalFormatValue
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### ContentTypes
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### Core
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### Custom
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### CustomProperty
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### DataValidation
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 24
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_error_style
      - set_input_title
      - ignore_blank
      - set_multi_range
      - allow_decimal_number
      - show_input_message
      - show_dropdown
      - allow_list_formula
      - allow_decimal_number_formula
      - allow_date_formula
      - allow_list_strings
      - allow_text_length_formula
      - allow_text_length
      - allow_time
      - set_error_message
      - set_input_message
      - show_error_message
      - allow_date
      - set_error_title
      - allow_any_value
      - allow_custom
      - allow_time_formula
      - allow_whole_number
      - allow_whole_number_formula
    #### ❌ Functions Not Yet Migrated
      - new
  ### Drawing
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 1
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - assemble_xml_file
    #### ❌ Functions Not Yet Migrated
      - new
  ### FeaturePropertyBag
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### FilterCondition
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 4
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - add_list_filter
      - add_custom_boolean_or
      - add_list_blanks_filter
      - add_custom_filter
    #### ❌ Functions Not Yet Migrated
      - new
  ### FilterData
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 2
    #### ❌ Functions Not Yet Migrated
      - new_string_and_criteria
      - new_number_and_criteria
  ### JsExcelData
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### JsExcelDataArray
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### JsExcelDataMatrix
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### Metadata
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 1
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - assemble_xml_file
    #### ❌ Functions Not Yet Migrated
      - new
  ### Packager
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### ProtectionOptions
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Functions Not Yet Migrated
      - new
  ### Relationship
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### RichString
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### RichValue
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### RichValueRel
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### RichValueStructure
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### RichValueTypes
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### Shape
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 10
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_width
      - set_alt_text
      - set_font
      - set_height
      - set_text_link
      - set_text_options
      - set_object_movement
      - set_format
      - set_url
      - set_text
    #### ❌ Functions Not Yet Migrated
      - textbox
  ### ShapeFont
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 11
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - unset_bold
      - set_color
      - set_name
      - set_strikethrough
      - set_italic
      - set_bold
      - set_size
      - set_pitch_family
      - set_underline
      - set_character_set
      - set_right_to_left
    #### ❌ Functions Not Yet Migrated
      - new
  ### ShapeFormat
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 6
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_solid_fill
      - set_no_fill
      - set_gradient_fill
      - set_no_line
      - set_pattern_fill
      - set_line
    #### ❌ Functions Not Yet Migrated
      - new
  ### ShapeGradientFill
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 3
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_type
      - set_angle
      - set_gradient_stops
    #### ❌ Functions Not Yet Migrated
      - new
  ### ShapeGradientStop
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Functions Not Yet Migrated
      - new
  ### ShapeLine
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 5
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_transparency
      - set_width
      - set_dash_type
      - set_color
      - set_hidden
    #### ❌ Functions Not Yet Migrated
      - new
  ### ShapePatternFill
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 3
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_background_color
      - set_foreground_color
      - set_pattern
    #### ❌ Functions Not Yet Migrated
      - new
  ### ShapeSolidFill
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 2
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_color
      - set_transparency
    #### ❌ Functions Not Yet Migrated
      - new
  ### ShapeText
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 3
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_direction
      - set_horizontal_alignment
      - set_vertical_alignment
    #### ❌ Functions Not Yet Migrated
      - new
  ### SharedStrings
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### SharedStringsTable
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### Sparkline
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 27
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - set_date_range
      - set_last_point_color
      - set_markers_color
      - set_right_to_left
      - set_column_order
      - set_high_point_color
      - set_line_weight
      - set_range
      - set_sparkline_color
      - set_type
      - show_high_point
      - show_negative_points
      - show_first_point
      - set_custom_max
      - set_group_min
      - set_style
      - set_first_point_color
      - show_hidden_data
      - show_low_point
      - set_group_max
      - set_low_point_color
      - set_negative_points_color
      - show_empty_cells_as
      - show_axis
      - show_markers
      - show_last_point
      - set_custom_min
    #### ❌ Functions Not Yet Migrated
      - new
  ### Styles
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### TableFunction
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### Theme
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 0
      - Migrated functions: 0
      - Not migrated functions: 0
  ### Vml
    #### Summary
      - Migrated methods: 0
      - Not migrated methods: 1
      - Migrated functions: 0
      - Not migrated functions: 1
    #### ❌ Methods Not Yet Migrated
      - assemble_xml_file
    #### ❌ Functions Not Yet Migrated
      - new
