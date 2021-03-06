#' @title ASAP Student Item Analysis Multi to Scantron Achievement Series
#' @description Convert the Student Item Analysis Multi export from Level 0 ASAP
#'   to the format of files exported from Scantron Achievement Series
#' @param asapFilePath Character of length 1 containing the filepath of an
#'   export from ASAP.
#' @param sectionTable data.frame showing associations between students and
#'   sections.  Columns include \code{StudentID}, \code{Course}, \code{Teacher},
#'   \code{Period}, and \code{Exam}. Each Exam value must match to exactly 1
#'   regents exam name using \code{grepl}.
#' @param merge character of length 1 indicating when to merge sections.
#' @param alphaResponses logical of length 1 - should MC responses be converted
#'   from digits to letters?  Defaults to TRUE.
#' @param messageLevel integer of length 1 indicating the level of messaging to
#'   print.  Defaults to \code{TP}.
#' @details For the \code{merge} parameter, the value will indicate whether to
#'   group different sections together.  The most grouped values are \code{T},
#'   \code{C}, or \code{P} (for grouping all students who have the same teacher,
#'   course, or period, respectively).  The least grouped value is \code{TCP}
#'   (for grouping students only when they have the exact same teacher, course,
#'   and period).  Other acceptable values are \code{TC}, \code{TP}, and
#'   \code{CP}.  Defaults to \code{TP}.
#'
#'   The export files generated by this function will be placed in a folder
#'   named \code{export} within the folder containing \code{asapFilePath}.
#' @return Nothing is returned by this function.
ASAP.SIAM_to_ScantronAS <- function(asapFilePath, sectionTable, merge = "TP", alphaResponses = T, additionalExams = NULL, messageLevel = 0) {
  datapath = paste0(dirname(asapFilePath),"/exports")
  if(!dir.exists(datapath)){dir.create(datapath)}                   # If the datafolder doesn't exist, create it

  asapData = read.csv(asapFilePath, stringsAsFactors = F, skip = 2) # Read in data exported from ASAP
  ItemLabels = unique(asapData$QuestionLabel)                       # Identify the item labels
  ItemNames = paste0("Q", ItemLabels)                               # Give them names by prepending "Q"
  Students = unique(asapData$textbox4)                              # Identify the set of students
  StudentFrame = strsplit(x = Students, split = ", ")               # Split the names
  StudentFrame = as.data.frame(StudentFrame, stringsAsFactors = F)  # Convert to a data.frame
  StudentFrame = as.data.frame(t(StudentFrame), stringsAsFactors = F)        # Transpose it
  rownames(StudentFrame) = NULL                                              # Eliminate row names
  colnames(StudentFrame) = c("Last", "First", "Student ID")                  # Add column names
  StudentFrame$Student = paste0(StudentFrame$Last, ", ", StudentFrame$First) # create full names
  StudentIDs = StudentFrame[,3]                                              # grab the field with the student IDs
  StudentIDs = regmatches(x = StudentIDs, m = regexpr(pattern = "^[[:digit:]]+", text = StudentIDs))         # limit it to just the actual ID
  StudentFrame$`Student ID` = StudentIDs                                                                     # put IDs back in the data.frame
  ExamName = unlist(read.csv(file = asapFilePath, stringsAsFactors = F, skip = 1, nrows = 1, header = F)[2]) # Get the exam name
  StudentFrame$`Test Name` = ExamName                                                                        # Add exam name to data.frame
  StudentFrame[,ItemNames] = ""                                                                              # Add columns for the items
  StudentFrame$Lookup = Students                                                                             # put raw student info in a column


  # Load per-student item response info
  for(stu in 1:nrow(StudentFrame)){
    thisStudent = StudentFrame$Lookup[stu]                    # identify the student
    StudentData = asapData[asapData$textbox4 == thisStudent,] # limit data to current student
    for(item in 1:length(ItemNames)){                         # for each item
      thisItem = ItemLabels[item]                             # identify the item
      response = StudentData$Response_1[StudentData$QuestionLabel == thisItem] # grab the response value

      # If it is a multiple choice question, convert the response to a letter
      if(!is.na(StudentData$CorrectResponseValue_1[StudentData$QuestionLabel == thisItem])){
        response = LETTERS[as.integer(response)]
      }

      StudentFrame[stu,ItemNames[item]] = response # load the response
    } # /for
  } # /for


  # Determine what exam this is and match it to the exam names in the sectionTable
  examNames = unique(sectionTable$Exam) # Load the set of exams
  if(!is.null(additionalExams)){
    examNames = unique(c(examNames, additionalExams))
  }
  isThisExam = logical(0)               # initialize logical vector to identify exam
  for(i in 1:length(examNames)){        # check each exam
    isThisExam = c(isThisExam, grepl(pattern = examNames[i], x = ExamName))
  }
  ExamName.official = examNames[isThisExam] # determine the official name of the exam

  # Error checking on exam name
  if(length(ExamName.official) == 0){
    stop(paste0("Did not recognize the exam ", ExamName))
  } else if(length(ExamName.official) > 1){
    stop(paste0("More than one match for ", ExamName, ", including ", paste0(examNames, collapse = ", ")))
  }

  # Load section info into test data
  sectionTable.current = sectionTable[sectionTable$Exam == ExamName.official,]             # Get a table of sections for this exam
  useRows = match(StudentFrame$`Student ID`, sectionTable.current$StudentID)               # Identify rows corresponding to students in export
  StudentFrame$Period = sectionTable.current$Period[useRows]                               # Load the period
  StudentFrame$Course = sectionTable.current$Course[useRows]                               # load the course name
  StudentFrame$Course = gsub(pattern = " ", replacement = "_", x = StudentFrame$Course)    # clean course names
  StudentFrame$Teacher = sectionTable.current$Teacher[useRows]                             # load the teacher name
  StudentFrame$Teacher = gsub(pattern = ", ", replacement = ".", x = StudentFrame$Teacher) # clean the teacher names

  # Create section names based on the desired merge
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
