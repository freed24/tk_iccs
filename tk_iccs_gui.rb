#!/usr/bin/ruby
# @Author: david
# @Date:   2017-07-04 18:16:06
# @Last Modified by:   anchen
# @Last Modified time: 2017-07-04 19:01:56
require 'tk'
require 'date'

# 图形界面生成
root = TkRoot.new { title "ICCS处理"}
frame_area_code = TkFrame.new(root){pack}
frame_quarter = TkFrame.new(root){pack}
frame_20 = TkFrame.new(root){pack}
frame_30 = TkFrame.new(root){pack}
frame_40 = TkFrame.new(root){pack}

TkLabel.new(frame_area_code){text '机构代码'
  pack(side:'left', padx:5, pady:10)}

Tk.mainloop