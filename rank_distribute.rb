# 
# Divid cells into groups according to the value of the cells currently selected in Excel
# assumptions:
#
# 	1, cells are selected in a running Excel application
# 	2, values are integers
#

require "win32ole"

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

class RankDistributor

	def distribute(data, seg_count, acending)
		# data should be an array of ranks(if value.to_i == 0 then ignore)
		ret = []
		@current_rank = nil

		seg_count.times do |s|
			ret << send(s+1 == seg_count ? :last_segment : :cal_segment, acending)
		end

		acending ? ret : ret.reverse
	end

	def optimal_distribute(data, seg_count)
		total = analyze(data)
		avg_group_size = (total / seg_count).to_i

		deltas = (0 .. avg_group_size * 5).to_a
		deltas = deltas.zip(deltas.map{|x| 0 - x }).flatten.uniq

		solutions = {}

		[true, false].each do |order|
			deltas.each do |d|
				@ideal_group_size = avg_group_size + d
				ret = distribute data, seg_count, order

				# puts "= Got solution:"

				# ret.each_with_index do |seg, idx|
				# 	puts "\tpoint: #{ret.count - idx}, count: #{seg.count}"
				# end

				groups_sizes = ret.map{|x| x.count}
				score = groups_sizes.standard_deviation

				solutions[score] = ret
			end
		end

		best_score = solutions.keys.min
		solutions[best_score]

	end

	def analyze(data)
		@rank_idx_map = {}

		data.each_with_index do |v, idx|
			t = v.to_i
			if t > 0
				@rank_idx_map[t] ||= []
				@rank_idx_map[t] << idx
			end
		end

		# b = @rank_idx_map.values.inject(0){|sum, x| sum + x.count}
		data.find_all{|x| x.to_i > 0}.count

	end

	def last_segment(acending = true)
		max_rank = @rank_idx_map.keys.max
		min_rank = @rank_idx_map.keys.min

		a = acending ? (@current_rank .. max_rank) : (min_rank .. @current_rank)
		ret = []

		a.to_a.each do |r|
			ret += @rank_idx_map[r] || []
		end

		ret
	end

	def cal_segment(acending = true)
		max_rank = @rank_idx_map.keys.max
		min_rank = @rank_idx_map.keys.min

		last_w_seg = @rank_idx_map.values.inject([]){|all, x| all + x}

		@current_rank ||= acending ? min_rank : max_rank

		if acending
			upper = @current_rank
			# incr upper rank until the last round is better than current;
			while upper <= max_rank

				seg = []
				(@current_rank ... upper).to_a.each do |r|
					seg += @rank_idx_map[r] || []
				end

				last_d = (last_w_seg.count - @ideal_group_size).abs
				curr_d = (seg.count - @ideal_group_size).abs

				# puts "last delta:#{last_d}, current delta:#{curr_d} (range: #{@current_rank} ... #{upper})"
				if (last_d > curr_d) || seg.empty?
					upper += 1
					last_w_seg = seg
				else
					@current_rank = upper - 1
					return last_w_seg
				end
			end
		else

			lower = @current_rank
			while lower >= min_rank

				seg = []
				arr = (lower .. @current_rank).to_a
				arr.shift

				arr.each do |r|
					seg += @rank_idx_map[r] || []
				end

				last_d = (last_w_seg.count - @ideal_group_size).abs
				curr_d = (seg.count - @ideal_group_size).abs

				if (last_d > curr_d) || seg.empty?
					lower -= 1
					last_w_seg = seg
				else
					@current_rank = lower + 1
					return last_w_seg
				end
			end
		end

		[]
	end

end

def scan_sheet(sheet, start_row_idx, start_col_idx, count, is_row = true)
	idx = is_row ? start_col_idx : start_row_idx

	count.times do 
		idx += 1
		c = is_row ? sheet.Cells(start_row_idx, idx) : sheet.Cells(idx, start_col_idx)
		yield(c) # true to stop
	end

end

if __FILE__ == $0
	
	# excel = WIN32OLE.new 'Excel.Application'
	# excel.Visible = 1
	# fn = ARGV[0]

	# wb = excel.Workbooks.Open File.join(File.dirname(__FILE__), fn)
	# sheet = wb.Worksheets 3
	excel = WIN32OLE.connect 'Excel.Application'
	wb = excel.ActiveWorkbook
	sheet = wb.ActiveSheet
	# puts sheet
	range = excel.Selection

	# puts range.Row
	# puts range.Column
	# puts range.Rows.Count
	# puts range.Columns.Count
	# abort

	row = range.Row
	col = range.Column
	data = []

	scan_sheet(sheet, row, col, range.Columns.Count) do |c|
		v = c.Value
		data << v unless v.nil?
	end

	d = RankDistributor.new
	ret = d.optimal_distribute data, 5

	puts "*" * 45
	puts "== Final solution =="

	ret.each_with_index do |seg, idx|
		v = 5 - idx
		puts "point: #{v}, count: #{seg.count}"

		seg.each do |c|
			sheet.Cells(row + 1, col + c).Value = v
		end
	end

	# wb.Save
	# wb.Close
	# excel.Quit
end