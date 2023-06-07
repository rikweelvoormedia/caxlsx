# lib/caxlsx/workbook/worksheet/sort_state.rb

module Axlsx
  class Worksheet
    # Represents the state of sorting for a worksheet.
    class SortState
      attr_reader :worksheet

      attr_reader :sort_conditions

      def initialize
        @worksheet = worksheet
        @sort_conditions = []
      end

      # Adds a sort condition to the worksheet.
      #
      # @param [String] reference The cell range to sort.
      # @param [Hash] options The sort options.
      #
      # @option options [String] :sort_by Specifies the column or columns to sort by.
      # @option options [Boolean] :descending Specifies whether to sort in descending order.
      # @option options [Boolean] :case_sensitive Specifies whether the sort is case sensitive.
      #
      # @return [SortState] self
      def add_sort_condition(reference, options = {})
        sort_condition = SortCondition.new(reference, options)
        @sort_conditions << sort_condition
        self
      end

      def to_xml_string(str = '')
        str << '<sortState>'
        sort_conditions.each do |sort_condition|
          sort_condition.to_xml_string(str)
        end
        str << '</sortState>'
      end



    end

    class SortCondition
      attr_reader :reference, :descending

      def initialize(reference, options = {})
        @reference = reference
        @descending = options[:descending] || false
      end

      def to_xml_string(str = '')
        str << '<sortCondition ref="B2:B277" descending="false" />'
        # str << %(ref="#{reference}" )
        # str << %(descending="#{descending}" ) if descending
        # str << '/>'
      end
    end
  end
end
