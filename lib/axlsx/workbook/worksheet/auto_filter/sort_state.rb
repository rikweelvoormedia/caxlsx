require 'axlsx/workbook/worksheet/auto_filter/sort_condition.rb'

module Axlsx
  # This class performs sorting on a range in a worksheet
  class SortState
   # creates a new SortState object
    # @param [AutoFilter] the auto_filter that this sort_state belongs to
    def initialize(auto_filter)
      @auto_filter = auto_filter

    end

    attr_reader :sort_conditions

    # A collection of SortConditions for this sort_state
    # @return [SimpleTypedList]
    def sort_conditions
      @sort_conditions ||= SimpleTypedList.new SortCondition
    end

    # Adds a SortCondition to the sort_state. This is the recommended way to add conditions to it.
    # It requires a col_id for the sorting, descending and the custom order are optional.
    # @param [Integer] col_id Zero-based index indicating the AutoFilter column to which the sorting should be applied to
    # @param [Boolean] a Boolean value if the order should be descending = true or false
    # @param [Array] An array containg a custom sorting list in order.
    # @return [SortCondition]
    def add_sort_condition(col_id, descending = false, custom_order = [])
      sort_conditions << SortCondition.new(col_id, descending, custom_order)
      sort_conditions.last
    end

    # method to increment the String representing the first cell of the range of the autofilter by 1 row for the sortCondition
    # xml string
    def increment_cell_value(str)
      letter = str[/[A-Za-z]+/]
      number = str[/\d+/].to_i

      incremented_number = number + 1

      "#{letter}#{incremented_number}"
    end

    # serialize the object
    # @return [String]
    def to_xml_string(str = +'')
      return unless sort_conditions != []

      ref = @auto_filter.range
      first_cell, last_cell = ref.split(':')
      ref = increment_cell_value(first_cell) + ':' + last_cell

      str << "<sortState xmlns:xlrd2='http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2' ref='#{ref}'>"
      sort_conditions.each { |sort_condition| sort_condition.to_xml_string(str, ref) }
      str << "</sortState>"
    end
  end
end
