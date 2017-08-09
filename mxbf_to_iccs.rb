#!/usr/bin/ruby
# @Author: anchen
# @Date:   2017-08-07 17:59:21
# @Last Modified by:   anchen
# @Last Modified time: 2017-08-07 18:03:19
require "win32ole"
require "dbf"

class MxbfToIccs
  AREA_CODE = '610100' # 机构代码
  INTERVAL = 0..46 # 表格数量
  MILLIONTH = 1/1000000.0 # 百万分之一，将原始数据由单位分转换为万元。
end

if __FILE__ == $0

end