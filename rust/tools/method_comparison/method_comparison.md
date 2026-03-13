# Method Comparison Report: rustxlsxwriter to wasm-xlsxwriter Migration

## Summary
  - ✅ Fully Migrated Structs: 48
  - ⚠️ Partially Migrated Structs: 13
  - ❌ Not Migrated Structs: 1
  - ✅ Migrated Enums: 47
  - ❌ Not Migrated Enums: 5
  - ✅ Total Migrated Methods: 606
  - ❌ Total Not Migrated Methods: 31
  - ✅ Total Migrated Functions: 70
  - ❌ Total Not Migrated Functions: 7
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
  ### ✅ Shape
  ### ✅ ShapeFont
  ### ✅ ShapeFormat
  ### ✅ ShapeGradientFill
  ### ✅ ShapeGradientStop
  ### ✅ ShapeLine
  ### ✅ ShapePatternFill
  ### ✅ ShapeSolidFill
  ### ✅ ShapeText
  ### ✅ Sparkline
  ### ✅ Table
  ### ✅ TableColumn
  ### ✅ Url
  ### ⚠️ Chart
    Summary
      - Migrated methods: 32
      - Not migrated methods: 2
      - Migrated functions: 10
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - add_series
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
      - Migrated methods: 14
      - Not migrated methods: 10
      - Migrated functions: 1
      - Not migrated functions: 0
    ❌ Methods Not Yet Migrated
      - allow_date
      - allow_date_formula
      - allow_decimal_number
      - allow_decimal_number_formula
      - allow_text_length
      - allow_text_length_formula
      - allow_time
      - allow_time_formula
      - allow_whole_number
      - allow_whole_number_formula
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
      - Migrated methods: 13
      - Not migrated methods: 0
      - Migrated functions: 1
      - Not migrated functions: 1
    ❌ Functions Not Yet Migrated
      - new_from_buffer
  ### ⚠️ Workbook
    Summary
      - Migrated methods: 10
      - Not migrated methods: 10
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
      - use_custom_theme
      - worksheets
      - worksheets_mut
  ### ⚠️ Worksheet
    Summary
      - Migrated methods: 145
      - Not migrated methods: 1
      - Migrated functions: 0
      - Not migrated functions: 2
    ❌ Methods Not Yet Migrated
      - add_conditional_format
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
