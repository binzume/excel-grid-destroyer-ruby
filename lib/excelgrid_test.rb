#!/usr/bin/ruby -Ku
# encoding: utf-8

require_relative 'excelgrid'

book = ExcelGrid::Book.new(Dir.glob("data/歌舞伎座13F_座席表*.xlsx").first)

s = book.sheet("sheet1")
p s.col('A1')

