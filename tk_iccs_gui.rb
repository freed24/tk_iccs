#!/usr/bin/ruby
# @Author: david
# @Date:   2017-07-04 18:16:06
# @Last Modified by:   anchen
# @Last Modified time: 2017-08-07 11:24:17
require 'tk'
require 'date'

class TkIccsGui
  def iccs_run

  end

  def initialize
    # 获取当前时间
    t = Time.now
    # 当前月份减一
    date = Date.new(t.year, t.month, t.day) << 1
    option = TkVariable.new('month')
    @year_month = Array.new

    # 业务数据处理
    create_data = proc{
      @area_code = @entry_code.get
      @dbf_dir = @entry_dbf_path.get + '/'
      @iccs_dir = @entry_iccs_path.get + '/*'
      if option.value == 'month'
        @year_month << @entry_year.get + @entry_month.get
      else
        month_list = case @entry_month.get
        when '1'
          %W[01 02 03]
        when '2'
          %W[04 05 06]
        when '3'
          %W[07 08 09]
        when '4'
          %W[10 11 12]
        end
        month_list.each{|month|
          @year_month << @entry_year.get + month
        }
      end
      iccs_run
    }

    # 图形界面生成
    root = TkRoot.new { title "ICCS处理"}
    frame_date = TkFrame.new(root){pack}
    frame_quarter = TkFrame.new(root){pack}
    frame_20 = TkFrame.new(root){pack}
    frame_30 = TkFrame.new(root){pack}
    frame_run = TkFrame.new(root){pack}

    TkLabel.new(frame_date){text '机构代码'
      pack(side:'left', padx:5, pady:10)}
    area_code = TkVariable.new
    area_code.value = '610100'
    @entry_code = TkEntry.new(frame_date, textvariable:area_code){
      width 6
      pack(side:'left', padx:5, pady:10)
    }

    TkLabel.new(frame_date){text '年度'
      pack(side: 'left', padx:5, pady:10)
    }
    year = TkVariable.new
    year.value = date.strftime("%Y")
    @entry_year = TkEntry.new(frame_date, textvariable:year){
      width 4
      pack(side: 'left', padx:2, pady:10)
    }

    label = TkLabel.new(frame_date){text '月份'
      pack(side:'left', padx:20, pady:10)
    }
    month = TkVariable.new
    month.value = date.strftime("%m")

    @entry_month = TkEntry.new(frame_date, textvariable:month){
      width 4
      pack(side:'left', padx:2, pady:10)
    }

    TkRadioButton.new(frame_quarter){
      text '月度'
      variable option
      value 'month'
      command{label.configure('text', '月份')
      month.value = date.strftime("%m")
      }
      pack side:'left'
    }

    # 季度判别
    quarter = case month.value.to_i
    when 1..3
      1
    when 4..6
      2
    when 7..9
      3
    when 10..12
      4
    end

    TkRadioButton.new(frame_quarter){
      text '季节'
      variable option
      value 'quarter'
      pack side:'left'
      command {label.configure('text', '季度')
        month.value = quarter
      }
    }

    # 数据目录选择
    #
    #
    # ICCS文件目录选择


    # 快报处理程序
    TkButton.new(frame_run){text '数据处理'
      width 10
      command create_data
      pack(side:'left', padx:10, pady:10)
    }

    TkButton.new(frame_run){text '退出'
      width 10
      command {exit}
      pack(side:'left', padx:20, pady:10)
    }
  end
end





TkIccsGui.new
Tk.mainloop