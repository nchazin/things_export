#get our output file name
set newFile to (choose file name with prompt "Output file:")
open for access newFile with write permission
set eof of newFile to 0

write "Item Name, Status, Due Date, Completion Date" & linefeed to newFile

tell application "Things"
	activate
	log completed now
	set AreaList to {"Inbox", "Today", "Next", "Scheduled", "Someday", "Projects"}
	repeat with AreaItem in AreaList
		write AreaItem & ":" & linefeed to newFile
		set toDoList to to dos of list AreaItem
		repeat with toDo in toDoList
			set todoName to name of toDo
			set todoCompletionDate to completion date of toDo
			set todoDueDate to due date of toDo
			set todoStatus to status of toDo
			#sub out for constants
			if todoStatus is open then
				set todoStatus to "open"
			else if todoStatus is completed then
				set todoStatus to "completed"
			end if
			
			if todoStatus is missing value then
				set todoStatus to "unknown"
			end if
			if todoCompletionDate is missing value then
				set todoCompletionDate to "''"
			else
				set todoCompletionDate to short date string of todoCompletionDate
			end if
			if todoDueDate is missing value then
				set todoDueDate to "''"
			else
				set todoDueDate to short date string of todoDueDate
			end if
			
			if (AreaItem as string = "Projects") then
				set projectToDoList to to dos of project todoName
				write " " & todoName & linefeed to newFile
				repeat with projectToDo in projectToDoList
					set prtodoName to name of projectToDo
					set prtodoCompletionDate to completion date of projectToDo
					set prtodoDueDate to due date of projectToDo
					set prtodoStatus to status of projectToDo
					#sub out for constants
					if prtodoStatus is open then
						set prtodoStatus to "open"
					end if
					if prtodoStatus is completed then
						set prtodoStatus to "completed"
					end if
					if prtodoCompletionDate is missing value then
						set prtodoCompletionDate to "''"
					else
						set prtodoCompletionDate to short date string of prtodoCompletionDate
					end if
					if prtodoDueDate is missing value then
						set prtodoDueDate to "''"
					else
						set prtodoDueDate to short date string of prtodoDueDate
					end if
					write prtodoName & "," & prtodoStatus & "," & prtodoDueDate & "," & prtodoCompletionDate & linefeed to newFile
				end repeat
				write linefeed to newFile
			else
				write todoName & "," & todoStatus & "," & todoDueDate & "," & todoCompletionDate & linefeed to newFile
			end if
			write linefeed to newFile
		end repeat
		write linefeed to newFile
	end repeat
end tell