#!/usr/bin/ruby
# @Author: david
# @Date:   2017-07-04 18:16:06
# @Last Modified by:   anchen
# @Last Modified time: 2017-08-04 18:44:15
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


    # 图形界面生成
    root = TkRoot.new { title "ICCS处理"}
    frame_area_code = TkFrame.new(root){pack}
    frame_quarter = TkFrame.new(root){pack}
    frame_20 = TkFrame.new(root){pack}
    frame_30 = TkFrame.new(root){pack}
    frame_40 = TkFrame.new(root){pack}

    TkLabel.new(frame_area_code){text '机构代码'
      pack(side:'left', padx:5, pady:10)}
    area_code = TkVariable.new
    area_code.value = '610100'
    @entry_code = TkEntry.new(frame_area_code, textvariable:area_code){
      width 6
      pack(side:'left', padx:5, pady:10)
    }

    TkLabel.new(frame_area_code){text '年度'
      pack(side: 'left', padx:5, pady:10)
    }
    year = TkVariable.new
    year.value = date.strftime("%Y")
    @entry_year = TkEntry.new(frame_area_code, textvariable:year){
      width 4
      pack(side: 'left', padx:2, pady:10)
    }

  end
end





TkIccsGui.new
Tk.mainloop