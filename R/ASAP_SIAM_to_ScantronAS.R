# library(openxlsx)
# library(dBtools)


ASAP.SIAM_to_ScantronAS <- function(asapFilePath, sectionTable, merge = "TP", messageLevel = 0) {
  datapath = paste0(dirname(asapFilePath),"/exports")
  if(!dir.exists(datapath)){dir.create(datapath)}                   # If the datafolder doesn't exist, create it

  asapData = read.csv(asapFilePath, stringsAsFactors = F, skip = 2)
  ItemLabels = unique(asapData$QuestionLabel)
  ItemNames = paste0("Q", ItemLabels)
  Students = unique(asapData$textbox4)
  StudentFrame = strsplit(x = Students, split = ", ")
  StudentFrame = as.data.frame(StudentFrame, stringsAsFactors = F)
  StudentFrame = as.data.frame(t(StudentFrame), stringsAsFactors = F)
  rownames(StudentFrame) = NULL
  colnames(StudentFrame) = c("Last", "First", "Student ID")
  StudentFrame$Student = paste0(StudentFrame$Last, ", ", StudentFrame$First)
  StudentIDs = StudentFrame[,3]
  StudentIDs = regmatches(x = StudentIDs, m = regexpr(pattern = "^[[:digit:]]+", text = StudentIDs))
  StudentFrame$`Student ID` = StudentIDs
  ExamName = unlist(read.csv(file = asapFilePath, stringsAsFactors = F, skip = 1, nrows = 1, header = F)[2])
  StudentFrame$`Test Name` = ExamName
  StudentFrame[,ItemNames] = ""
  StudentFrame$Lookup = Students


  for(stu in 1:nrow(StudentFrame)){
    thisStudent = StudentFrame$Lookup[stu]
    StudentData = asapData[asapData$textbox4 == thisStudent,]
    for(item in 1:length(ItemNames)){
      thisItem = ItemLabels[item]
      StudentFrame[stu,ItemNames[item]] = StudentData$Response_1[StudentData$QuestionLabel == thisItem]
    } # /for
  } # /for



  examNames = unique(sectionTable$Exam)
  isThisExam = logical(0)
  for(i in 1:length(examNames)){
    isThisExam = c(isThisExam, grepl(pattern = examNames[i], x = ExamName))
  }
  ExamName.official = examNames[isThisExam]

  if(length(ExamName.official) == 0){
    stop(paste0("Did not recognize the exam ", ExamName))
  } else if(length(ExamName.official) > 1){
    stop(paste0("More than one match for ", ExamName, ", including ", paste0(examNames, collapse = ", ")))
  }


  sectionTable.current = sectionTable[sectionTable$Exam == ExamName.official,]
  useRows = match(StudentFrame$`Student ID`, sectionTable.current$StudentID)
  StudentFrame$Period = sectionTable.current$Period[useRows]
  StudentFrame$Course = sectionTable.current$Course[useRows]
  StudentFrame$Course = gsub(pattern = " ", replacement = "_", x = StudentFrame$Course)
  StudentFrame$Teacher = sectionTable.current$Teacher[useRows]
  StudentFrame$Teacher = gsub(pattern = ", ", replacement = ".", x = StudentFrame$Teacher)


  if(merge == "TP"){
    StudentFrame$Section = paste0(StudentFrame$Teacher, "_p", StudentFrame$Period)
  } else if(merge == "TCP"){
    StudentFrame$Section = paste0(StudentFrame$Teacher, "_", StudentFrame$Course, "_p", StudentFrame$Period)
  } else if(merge == "P"){
    StudentFrame$Section = paste0("p", StudentFrame$Period)
  } else if(merge == "TC"){
    StudentFrame$Section = paste0(StudentFrame$Teacher, "_", StudentFrame$Course)
  }else if(merge == "CP"){
    StudentFrame$Section = paste0(StudentFrame$Course, "_p", StudentFrame$Period)
  } else if(merge == "T"){
    StudentFrame$Section = StudentFrame$Teacher
  } else if(merge == "C"){
    StudentFrame$Section = StudentFrame$Course
  } else {
    stop(paste0("The merge code ", merge, " was not recognized."))
  }


  sections = unique(StudentFrame$Section)
  OutputList = vector(mode = "list", length = length(sections))
  names(OutputList) = sections

  StudentFrame = StudentFrame[order(StudentFrame$Last, StudentFrame$First, StudentFrame$`Student ID`),]

  cols2use = colnames(StudentFrame)
  cols2use = c("Student", cols2use[!(cols2use %in% c("Last", "First", "Student", "Lookup", "Period"))])
  StudentFrame = StudentFrame[,cols2use]

  for(thisSection in 1:length(sections)){
    currentSection = sections[thisSection]
    currentStudentFrame = StudentFrame[StudentFrame$Section == currentSection,]
    row.names(currentStudentFrame) = NULL
    currentStudentFrame[nrow(currentStudentFrame) + 1,] = " "
    OutputList[[thisSection]] = currentStudentFrame
  }

  for(thisSection in 1:length(sections)){
    write.csv(x = OutputList[[thisSection]], file = paste0(datapath, "/", names(OutputList)[thisSection], "_itemresponses.csv"), row.names = F)
  }




} # /function
