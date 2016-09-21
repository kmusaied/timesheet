Sys.setlocale(local = "Arabic_Saudi Arabia.1256")

require(openxlsx)
require(dplyr)
require(rpivotTable)

x <- NULL 

run <- function(month, week) {
  week = paste("W", week , sep = "")
  folder <- paste("TimeSheets", month, week , sep = "/")
  
  files <- list.files(path = folder)
  
  for (f in files) {
    filePath <- paste(folder, f, sep = "/")
    d <-  read.xlsx(filePath, detectDates = TRUE)
    x <<- merge(x, d, all = TRUE)
  }
  colnames(x) <<-
    c(
      "EmpName",
      "Task",
      "System",
      "CR_Proj",
      "TFSID",
      "Task_Type",
      "Date",
      "WeekDay",
      "Hours"
    )
  GroupByEmp <<- aggregate(Hours ~ EmpName  , x , sum)
  GroupByWeekDay <<-
    aggregate(Hours ~ EmpName + WeekDay  , x , sum)
  GroupByType <<-
    aggregate(Hours ~ EmpName + Task_Type  , x , sum)
  GroupByDate <<- aggregate(Hours ~ Date  , x , sum)
  AvaragePerHours <<- aggregate(Hours ~ EmpName   , x , mean)
  TasksPerProject <<- aggregate(EmpName ~ CR_Proj   , x , length)
  HoursPerDay <<- aggregate(Hours ~ Date , x , sum)
  #tasks <<- x[,c(x$EmpName)]
  
  pie(
    GroupByEmp$Hours,
    labels = GroupByEmp$EmpName,
    col = rainbow(length(GroupByEmp$EmpName)),
    main = "Total Hours per Employee"
  )
  
} 

chart <- function(c)
{
  switch (
    c,
    "1" = pie(
      GroupByEmp$Hours,
      labels = GroupByEmp$EmpName,
      col = rainbow(length(GroupByEmp$EmpName)),
      main = "Total Hours per Employee"
    )
    ,
    "2" = pie(
      AvaragePerHours$Hours,
      labels = AvaragePerHours$EmpName,
      col = rainbow(length(AvaragePerHours$EmpName)),
      main = "Avarage Hours per Employee"
    )
    ,
    "3" = pie(
      TasksPerProject$EmpName,
      labels = TasksPerProject$CR_Proj,
      col = rainbow(length(TasksPerProject$CR_Proj)),
      main = "Tasks per Project"
    )
    ,
    "4" = {
      pie(
        HoursPerDay$Hours,
        labels = HoursPerDay$Hours,
        col = rainbow(length(HoursPerDay$Date)),
        main = "Hours Per Day"
      )
      legend("topright",
             legend =  HoursPerDay$Date,
             fill = rainbow(length(HoursPerDay$Date)))
    },
    "5" = {
      rpivotTable(
        x[1:9],
        rows = c("EmpName", "Task_Type") ,
        col = "Date",
        vals = "Hours",
        aggregatorName = "Sum"
      )
    }
    
  )
  
}
