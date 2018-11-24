Attribute VB_Name = "DAOModule"
Option Explicit

Public Sub getExcelData()
  Dim db As DAO.Database
  Dim rs As DAO.Recordset
  Dim calendarData As String
  calendarData = "C:\Users\mike\Documents\My Projects\Project.Scheduler\calendar.xls"
  Set db = OpenDatabase(Name:=calendarData, Options:=False, ReadOnly:=False, Connect:="Excel 8.0")
  Set rs = db.OpenRecordset("SELECT * FROM `jobSchedule` WHERE Subject <> null")
  
  While Not rs.EOF
      Debug.Print rs.Fields(6).Value
      rs.MoveNext
  Wend
  
  rs.Close
  db.Close
  
  Set rs = Nothing
  Set db = Nothing
End Sub
