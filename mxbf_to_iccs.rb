#!/usr/bin/ruby
# @Author: anchen
# @Date:   2017-08-07 17:59:21
# @Last Modified by:   anchen
# @Last Modified time: 2017-08-09 17:20:32
require "win32ole"
require "dbf"

class MxbfToIccs
  AREA_CODE = '610100' # 机构代码
  INTERVAL = 0..46 # 表格数量
  MILLIONTH = 1/1000000.0 # 百万分之一，将原始数据由单位分转换为万元。

  def initialize(date_source: dbf_dir, out_data_dir: xls_dir, year: '2017', month: '??', day: '??', cc: '??')
    @date_source = date_source
    @out_data_dir = out_data_dir
    @year, @month, @day, @cc = year, month, day, cc
    @quarter = "#{@year}#{@month}"
    @rh_guoKu_47 = %W[0011001 1011001 2011001 3011001 4011001 5011001 6011001 7011001 8011001] # 47 人行国库处 01110000001100
    @xh_zhouZhi_1 = %W[0402211]   # 周至县农村信用合作联社 C3250761000015
    @xh_lianShe_2 = %W[0402150]   # 陕西省农村信用社联合社 C3250661000014
    @xh_xianYang_3 = %W[0402280]   # 咸阳市渭城区农村信用合作联社 C3181361000014
    @xh_linTong_4 = %W[0402171] # 西安市临潼区农村信用合作联社
    @xh_huXian_5 = %W[0402191 0402192] # 户县农村信用合作联社
    @xh_gaoLing_6 = %W[0402231] # 高陵县农村信用合作联社
    @xh_lanTian_7 = %W[0402271 0402272] # 蓝田县农村信用合作联社(2个交换行) C3178261000017
    @qn_lianHu_8 = %W[0412071 0412072 0412087 0412088 0412079 0412082 0412078 0412077 0412086 0412081 0412075 0412073 0412074 0412076 0412083 0412080 0412085] # 西安市莲湖区农村信用合作联社（交行行数量：15） C3178161000016
    @qn_baQiao_9 = %W[0412023 0412091 0412092 0412093 0412094 0412095 0412096 0412097 0412098 0412099] #  西安市灞桥区农村信用合作联社（交行行数量：6） C3178061000015
    @qn_xinCheng_10 = %W[0412131] # 西安市新城区农村信用合作联社 C3177961000012
    @xh_yanLiang_11 = %W[0402251] # 西安市阎良区农村信用合作联社 C3177861000011
    @xh_changAn_12 = %W[0402111 0402112 0402113 0402114 0402115 0402116 0402117 0402118 0402119 0402120 0402121 0402122 0402123 0402366] # 西安市长安区农村信用合作联社（交行行数量：12） C3177761000010
    @qn_weiYang_13 = %W[0412047 0412048 0412031 0412024 0412026 0412043 0412033 0412036 0412035 0412041 0412034 0412062 0412037 0412025
    0412032 0412049 0412038 0412039 0412045 0412040 0412044 0412042 0412046] #  西安市未央区农村信用合作联社（交行行数量：20）
    @qn_YanTa_14 = %W[0412051 0412058 0412054 0412355 0412353 0412362 0412055 0412354 0412052 0412356 0412352 0412358 0412057 0412063 0412064 0412065
    0412364 0412056 0412360 0412359 0412059 0412363 0412053 0412365 0412060 0412361 0412061 0412369 0412066 0412067 0412367 0412357 0412371 0412372 0412373] # 西安市雁塔区农村信用合作联社 （交行行数量：25）
    @qn_beiLin_15 = %W[0412009 0412011 0412015 0412018 0412021 0412020 0412017 0412013 0412016 0412022 0412012 0412014 0412019 0412368] #  西安市碑林区农村信用合作联社（13个交换行） C3177461000017
    @zlbs = 0; @zlje = 0
    @jfje = 0; @dfbs = 0; @dfje = 0; @hjje = 0
    @interval = 0..46
    row = INTERVAL.max + 1 # 表格行数
    @zb_data = Array.new(row) { Array.new(2*row, 0) }
    @fb_data = Array.new(row) { Array.new(14, 0) }

    @excel = WIN32OLE.new('excel.Application')

    @gonghang_jf_je = 0; @gonghang_df_je = 0
  end

  # 明细备份数据统计
  def data_statistics()
    files = Dir[@data_source + "#{@year}#{@month}#{@day}#{@cc}mxbf.dbf"]
    files.each do |file|
      mxbf_table = DBF::Table.new(file)
      mxbf_table.each do |record|
        next if record.trhh.size != 7 || record.jym == 8 || record.jym == 28
        @trhh = record.trhh
        @tchh = record.tchh
        @jym  = record.jym
        @je   = record.je
        debit_credit_bill() # 根据票据借贷方向确定数据行列
      end
    end
  end

  # 获取表格文件列表
  def data_write_xls_files()
    jhh_number = jhh_count()
    fb_row = INTERVAL.max
    INTERVAL.each do |t|
      @zlbs += (@fb_data[t][10] + @fb_data[t][8])
      @zlje += ((@fb_data[t][11] + @fb_data[t][9])*MILLIONTH).round(2)
      INTERVAL.each do |e|
        @zb_data[t][2*e+1] = (@zb_data[t][2*e+1]*MILLIONTH).round(2) if @zb_data[t][2*e+1] != 0
      end
    end

    Dir.glob("#{@out_data_dir}/*.xls") do |file|
      work = @excel.Workbooks.Open(File.expand_path(file))
      @sheet=work.Worksheets(1)
      case File.basename(file)
      when 'CJBD-LLLX-TC-IN(同城流量流向-行业间).xls'
        @sheet.Range("B3").value = @quarter
        @sheet.Range("D7:AS27").value = 0
      when 'CJBD-XT-TC(同城清算系统指标采集表单)_ICCS0000610100_月_610100_市.xls'
        @sheet.Range("B2").value = @quarter
        @sheet.Range("F2").value = AREA_CODE
        @sheet.Range("G4, G6").value = jhh_number, jhh_number
        @sheet.Range("G9").value = jhh_number - 9 # 交换行数量减去9家人行国库
        @sheet.Range("G12:H12").value = @zlbs, @zlje
      when 'CJBD-LLLX-TC-ON(同城流量流向-机构间).xls'
        @sheet.Range("B3").value = @quarter
        @sheet.Range("D7:CS53").value = @zb_data
      else
        # 数据行数与表格序号对应
        data_write(fb_row); fb_row -= 1
      end
      work.Close(1)
    end
    @excel.Quit()
  end

  # 将数据写入46个分表
  def data_write(table_order_number)
    t=table_order_number
    @sheet.Range("B2").value = @quarter   # 期数
    @sheet.Range("F2").value = AREA_CODE  # 地区代码
    @sheet.Range("G5:H5").value = @fb_data[t][0], (@fb_data[t][1]*MILLIONTH).round(2)
    @sheet.Range("G6:H6").value = @fb_data[t][2], (@fb_data[t][3]*MILLIONTH).round(2)
    @sheet.Range("G7:H7").value = @fb_data[t][4], (@fb_data[t][5]*MILLIONTH).round(2)
    @sheet.Range("G8:H8").value = 'NAP'
    @sheet.Range("G12:H12").value = @fb_data[t][6], (@fb_data[t][7]*MILLIONTH).round(2)
    @sheet.Range("G14:H14").value = @fb_data[t][10], (@fb_data[t][11]*MILLIONTH).round(2)
    @sheet.Range("G15:H15").value = @fb_data[t][8], (@fb_data[t][9]*MILLIONTH).round(2)
    @sheet.Range("G4").value = @fb_data[t][0] + @fb_data[t][2] + @fb_data[t][4] + @fb_data[t][6]
    @sheet.Range("G13").value = @fb_data[t][8] + @fb_data[t][10]
    @sheet.Range("H4").value = ((@fb_data[t][1] + @fb_data[t][3] + @fb_data[t][5]+ @fb_data[t][7])*MILLIONTH).round(2)
    @sheet.Range("H13").value = ((@fb_data[t][9] + @fb_data[t][11])*MILLION).round(2)
  end

   private
  # 按交换行分类
  def bill_sort(jhh, i=1)
    case
    when @xh_zhouZhi_1.include?(jhh)
      l = 0 * i
    when @xh_lianShe_2.include?(jhh)
      l = 1 * i
    when @xh_xianYang_3.include?(jhh)
      l = 2 * i
    when @xh_linTong_4.include?(jhh)
      l = 3 * i
    when @xh_huXian_5.include?(jhh)
      l = 4 * i
    when @xh_gaoLing_6.include?(jhh)
      l = 5 * i
    when @xh_lanTian_7.include?(jhh)
      l = 6 * i
    when @qn_lianHu_8.include?(jhh)
      l = 7 * i
    when @qn_baQiao_9.include?(jhh)
      l = 8 * i
    when @qn_xinCheng_10.include?(jhh)
      l = 9 * i
    when @xh_yanLiang_11.include?(jhh)
      l = 10 * i
    when @xh_changAn_12.include?(jhh)
      l = 11 * i
    when @qn_weiYang_13.include?(jhh)
      l = 12 * i
    when @qn_YanTa_14.include?(jhh)
      l = 13 * i
    when @qn_beiLin_15.include?(jhh)
      l = 14 * i
    when /^0319/.match(jhh) # 昆仑银行股份有限公司（0319）16
      l = 15 * i
    when /^0320/.match(jhh) # 重庆银行 （0320）17
      l = 16 * i
    when /^0313/.match(jhh) # 西安银行股份有限公司 （0313）18
      l = 17 * i
    when /^0314/.match(jhh) # 长安银行股份有限公司 （0314）19
      l = 18 * i
    when /^0317/.match(jhh) # 齐商银行 （0317） 20
      l = 19 * i
    when /^0318/.match(jhh) # 宁夏银行 （0318） 21
      l = 20 * i
    when /^0315/.match(jhh) # 北京银行 （0315） 22
      l = 21 * i
    when /^0330/.match(jhh) # 成都银行 （0330） 23
      l = 22 * i
    when /^0503/.match(jhh) # 渣打银行 （0503） 24
      l = 23 * i
    when /^0597/.match(jhh) # 韩亚银行 （0597） 25
      l = 24 * i
    when /^0501/.match(jhh) # 东亚银行 （0501） 26
      l = 25 * i
    when /^0502/.match(jhh) # 汇丰银行 （0502） 27
      l = 26 * i
    when /^0403/.match(jhh) # 邮储银行 （0403） 28
      l = 27 * i
    when /^0316/.match(jhh) # 浙商银行 （0316） 29
      l = 28 * i
    when /^0312/.match(jhh) # 恒丰银行 （0312） 30
      l = 29 * i
    when /^0310/.match(jhh) # 浦发银行 （0310） 31
      l = 30 * i
    when /^0309/.match(jhh) # 兴业银行 （0309） 32
      l = 31 * i
    when /^0308/.match(jhh) # 招商银行 （0308） 33
      l = 32 * i
    when /^0307/.match(jhh) # 平安银行 （0307） 34
      l = 33 * i
    when /^0305/.match(jhh) # 民生银行 （0305） 35
      l = 34 * i
    when /^0311/.match(jhh) # 华夏银行 （0311） 36
      l = 35 * i
    when /^0303/.match(jhh) # 光大银行 （0303） 37
      l = 36 * i
    when /^0302/.match(jhh) # 中信银行 （0302） 38
      l = 37 * i
    when /^0301/.match(jhh) # 交通银行 （0301） 39
      l = 38 * i
    when /^0203/.match(jhh) # 中国农业发展银行 （0203） 40
      l = 39 * i
    when /^0202/.match(jhh) # 中国进出口银行 （0202） 41
      l = 40 * i
    when /^0201/.match(jhh) # 国家开发银行 （0201） 42
      l = 41 * i
    when /^0105/.match(jhh) # 建设银行 （0105） 43
      l = 42 * i
    when /^0104/.match(jhh) # 中国银行 （0104） 44
      l = 43 * i
    when /^0103/.match(jhh) # 农业银行 （0103） 45
      l = 44 * i
    when /^0102/.match(jhh) # 工商银行 （0102） 46
      l = 45 * i
    when @rh_guoKu_47.include?(jhh) # 人行国库处 47
      l = 46 * i
    when /^0306/.match(jhh) # 广发银行 （0306） *24
      l = 24 * i
    when /^0350/.match(jhh) # 渤海银行 （0350） *24
      l = 24 * i
    end
  end

  # 票据分类
  def bill_jym_sort()
    case @jym
    when 1,21,51,71 #支票
      fl = 0
    when 2,22 #本票
      fl = 2
    when 3,23 #汇票
      fl = 4
    when 8,28 #非清算票据
      fl = 12
    else
      fl = 6 # 其它类
    end
    return fl
  end

  # 票据借贷分类
  def debit_credit_bill()
    if @jym < 50 then
      c = bill_sort(@trhh); r = bill_sort(@tchh, i=2)
      @fb_data[c][8] += 1
      @fb_data[c][9] += @je
      @fb_data[r/2][10] += 1
      @fb_data[r/2][11] += @je
    else
      @dfbs += 1; @dfje += @je
      c = bill_sort(@tchh); r = bill_sort(@trhh, 2)
      @fb_data[r/2][10] += 1
      @fb_data[r/2][11] += @je
      @fb_data[c][8] += 1
      @fb_data[c][9] += @je
    end
    @zb_data[c][r] += 1
    @zb_data[c][r+1] += @je
    fl = bill_jym_sort() # 根据交易码类型获取票据种类
    @fb_data[c][fl] += 1
    @fb_data[c][fl+1] += @je
  end

end

if __FILE__ == $0

end