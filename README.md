# VBA
'相对引用--绝对引用
for 循环
  if 条件 then
    exit for
  end if 
next
range("A2")
range("A" & i)'变量引用
worksheets'工作表对象(方法：Select Add Delete Copy 属性：Count Name)
eg: worksheets.Add after:=Worksheets(Worksheets.count)'工作表后插入表格
Application
eg： Worksheet1.select = sheets(1月).select = sheets(1).select
Application.DisplayAlerts = False(True)'显示警告，恢复
