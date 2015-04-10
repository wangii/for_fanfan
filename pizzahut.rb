#
# Pizzahut 241 code generator (UK)
# run the code, then use the code in http://www.tellpizzahut.co.uk to generate
# a 2 for 1 voucher
#

def gen_code
    ret     = '1' * 14
    d       = Time.now - 3600 * 24 * 3

    ret[6]  = d.strftime('%m')[0]
    ret[7]  = d.strftime('%m')[1]

    ret[0]  = d.strftime('%d')[0]
    ret[1]  = d.strftime('%d')[1]

    ret[9]  = d.strftime('%y')[0]
    ret[10] = d.strftime('%y')[1]

    [2,3,4,5,8,11,12].each do |idx|
        ret[idx] = rand(10).to_s
    end

    tmp     = ret.split(//)
    tmp.pop

    counter = tmp.length

    while counter > -1 do
        tmp[counter - 1] = tmp[counter - 1].to_i * 2
        tmp[counter - 2] = tmp[counter - 2].to_i
        counter -= 2
    end

    tmp = tmp.map{|x| x.to_s}.join

    sum = 0
    tmp.split(//).each do |idx|
        sum += tmp[idx].to_i
    end

    ret[13] = (sum * 9 % 10).to_s
    ret
end

puts gen_code
