# 
# Divid cells into groups according to the value of the cells currently selected in Excel
# assumptions:
#
# 	1, cells are selected in a running Excel application
# 	2, values are integers
# 	3, write to cells below the current row

require "win32ole"

# Basic stats methods
module Enumerable

    def sum
      self.inject(0){|accum, i| accum + i }
    end

    def mean
      self.sum/self.length.to_f
    end

    def sample_variance
      m = self.mean
      sum = self.inject(0){|accum, i| accum +(i-m)**2 }
      sum/(self.length - 1).to_f
    end

    def standard_deviation
      return Math.sqrt(self.sample_variance)
    end
end 

DataPoint = Struct.new(:rank, :pos, :tag) do
	def with_extra
		@rank + @extra
	end
end

# Collection of data points
class Segment < Array
	attr_accessor :left, :right

	def left_mutate
		return [clone, []] if map{|x| x.rank}.min > left
		cp = []
		out = []

		each do |x|
		end
	end

	def right_mutate
		return [clone, []] if map{|x| x.rank}.max < right
	end
end

class Solution
	attr_accessor :segments

	def mutate(idx, mutate_downward)
	end

	def score
		@score ||= @segments.map{|x| x.count}.standard_deviation
	end

	def <<(seg)
		@segments ||= []
		@segments << seg
	end

	def has_empty_segment?
		@groups.find_all{|x| x.empty? }.count > 0
	end

	def tail_empty?
		@segments.first.empty? || @segments.last.empty?
	end
end

module Distributer
	
	def self.distribute(data, seg_count)
		if data.length == data.map{|x| x.rank}.uniq.length && data.length >= seg_count
			SimpleDistributor.distribute data, seg_count
		end
	end

	module SimpleDistributor
		def self.distribute(data, seg_count)
			return [data] if seg_count == 1

			@sorted = data.sort{|a, b| a.rank <=> b.rank}
			head_count = @sorted.length/seg_count

			head = @sorted[0...head_count]
			tail = @sorted[head_count .. -1]
			[head] + distribute(tail, seg_count - 1)
		end
	end
end

def scan_sheet(sheet, start_row_idx, start_col_idx, count, is_row = true)
	idx = is_row ? start_col_idx : start_row_idx

	count.times do |t|
		c = is_row ? sheet.Cells(start_row_idx, idx + t) : sheet.Cells(idx + t, start_col_idx)
		yield(c) # true to stop
	end
end

if __FILE__ == $0
	
	excel = WIN32OLE.connect 'Excel.Application'
	wb = excel.ActiveWorkbook
	sheet = wb.ActiveSheet
	range = excel.Selection

	row = range.Row
	col = range.Column

	is_row = range.Rows.count == 1
	data = []

	scan_sheet(sheet, row, col, (is_row ? range.Columns : range.Rows).Count, is_row) do |c|
		v = c.Value.to_i
		unless v == 0
			p = DataPoint.new

			p.rank = v
			p.pos = [c.Row, c.Column]

			data << p
		end
	end

	n = 5

	Distributer.distribute(data, n).each_with_index do |seg, idx|
		v = n - idx
		puts "point: #{v}, count: #{seg.count}"

		seg.each do |r|
			x, y = r.pos
			is_row ? x+=1 : y+=1

			sheet.Cells(x, y).Value = v
		end
	end
end