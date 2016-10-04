rm(list=ls(all=TRUE))

Sys.setlocale(local = "Arabic_Saudi Arabia.1256")

require(openxlsx)
require(dplyr)
require(rpivotTable)


x<<- NULL

run <- function(month, week) {
  week = paste("W", week , sep = "")
  folder <<- paste("TimeSheets", month, week , sep = "/")
  
  files <- list.files(path = folder)
  x<<- NULL
  for (f in files) {
    message(f)
    filePath <- paste(folder, f, sep = "/")
    d <-  read.xlsx(filePath, detectDates = TRUE)
    x <<- rbind( x,d)
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
      "Hours",
      "Notes"  
    )
  
  x2 <<- filter(x,Task_Type !="إجازة"  )
  x2 <<- filter(x2,Task_Type !="إستئذان")
  
  GroupByEmp <<- aggregate(Hours ~ EmpName  , x2 , sum)
  GroupByWeekDay <<-
    aggregate(Hours ~ EmpName + Date  , x2 , sum)
  GroupByType <<-
    aggregate(Hours ~ EmpName + Task_Type  , x2 , sum)
  GroupByDate <<- aggregate(Hours ~ Date  , x2 , sum)
  AvaragePerHours <<- aggregate(Hours ~ EmpName   , x2 , mean)
  TasksPerProject <<- aggregate(EmpName ~ CR_Proj   , x2 , length)
  TasksPerSystem <<- aggregate(EmpName ~ System   , x2 , length)
  HoursPerDay <<- aggregate(Hours ~ Date , x2 , sum)
  Over8HoursPerDay <<- filter(GroupByWeekDay, Hours > 7)
  #tasks <<- x[,c(x$EmpName)]
  
  pie(
    GroupByEmp$Hours,
    labels = GroupByEmp$Hours,
    col = rainbow(length(GroupByEmp$EmpName)),
    main = "Total Hours per Employee"
  )
  legend("bottomleft",
         legend =  GroupByEmp$EmpName, cex = 0.7,
         fill = rainbow(length(GroupByEmp$EmpName)))
  
  
  
} 

Validate <- function ()
{
  if (is.null(x))
  {
    message("Please call Run function first.")
    returnValue()
  }
  
  #Check Task Types
  InvalidTaskTypes <-  filter( x2 ,  !Task_Type %in% 
             c("مشروع"
               ,"تشغيل"
               ,"دراسة"
              ,"إستئذان"
               ,"إجازة"
               ))
  
  
  
  
}

ValidateDates <- function (startDate , endDate)
{
  if (is.null(x))
  {
    message("Please call Run function first.")
    returnValue()
  }
  
  invalid_Dates <<-  filter(x, Date < StartDate , Date > endDate)
}


#Generate charts 
chart <- function(c)
{
  if (is.null(x))
  {
    message("Please call Run function first.")
    returnValue()
  }
  
  switch (
    c,
    "0" = {
      rpivotTable(x[1:9])
    },
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
    ,
    "6" = {
      rpivotTable(
        x[1:9],
        rows = c("EmpName", "Task") ,
        vals = "Hours",
        aggregatorName = "Sum"
      )
    }
    , "7" = {
      x$Task_Type <- as.factor(x$Task_Type)
      for (tt in levels(x$Task_Type)) {
        
        pie(x[,x$Task_Type==tt])
        
        
      }
      
    }
  )
  
}
Export = function(month , week)
{
  run(month,week)
  fileName <- paste("timeSheet",month,"W",week,".xlsx", sep = "_")
  write.xlsx(x, fileName)
}