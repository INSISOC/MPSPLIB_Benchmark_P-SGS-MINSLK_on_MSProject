Attribute VB_Name = "Rescheduling_Module"
Option Explicit


Sub ResourceLevelingEngine()

    ' === Variable declarations ===
    Dim Project As Project
    Dim startDate As Date, endDate As Date
    Dim currentDate As Date, newDate As Date, limitDate As Date
    Dim macroStartTime As Double, macroEndTime As Double, macroElapsedTime As Double

    Dim candidateTasks As Collection, ongoingTasks As Collection, completedTasks As Collection
    Dim task As task, tempTask As task, t As task, candidate As task
    Dim i As Integer, j As Integer, k As Integer, m As Integer

    Dim resourceCount As Integer
    Dim availableUnits() As Double
    Dim res As Resource
    Dim a As Assignment
    Dim overallocatedCount As Double, totalUnits As Double

    ' === Initialization ===
    Set Project = ActiveProject
    startDate = Project.ProjectStart
    endDate = Project.ProjectFinish
    currentDate = startDate
    macroStartTime = Timer

    ' === Main loop ===
    While DetectOverallocation = True
    
        ' === Identify completed tasks ===
        Set completedTasks = New Collection
        For Each task In Project.Tasks
            If Not task Is Nothing Then
                If task.Finish <= currentDate Then
                    completedTasks.Add task
                End If
            End If
        Next task

        ' === Identify ongoing tasks ===
        Set ongoingTasks = New Collection
        For Each task In Project.Tasks
            If Not task Is Nothing Then
                If task.Start < currentDate And task.Finish > currentDate Then
                    ongoingTasks.Add task
                End If
            End If
        Next task

        ' === Identify candidate tasks ===
        Set candidateTasks = New Collection
        For Each task In Project.Tasks
            If Not task Is Nothing Then
                If Left(task.Start, 10) = Left(currentDate, 10) Then
                    candidateTasks.Add task
                End If
            End If
        Next task

        ' === Sort candidate tasks by Total Slack (simple bubble sort) ===
        For i = 1 To candidateTasks.Count - 1
            For j = 1 To candidateTasks.Count - 1
                If candidateTasks(j).TotalSlack > candidateTasks(j + 1).TotalSlack Then
                    Set tempTask = candidateTasks(j)
                    candidateTasks.Remove j
                    candidateTasks.Add tempTask
                End If
            Next j
        Next i

        ' === Resource availability ===
        resourceCount = Project.resourceCount
        ReDim availableUnits(1 To resourceCount)

        k = 1
        For Each res In Project.Resources
            If Not res Is Nothing Then
                availableUnits(k) = res.MaxUnits
                k = k + 1
            End If
        Next res

        ' === Subtract usage from ongoing tasks ===
        For Each t In ongoingTasks
            k = 1
            For Each res In Project.Resources
                If Not res Is Nothing Then
                    For Each a In t.Assignments
                        If a.Resource.ID = res.ID Then
                            availableUnits(k) = availableUnits(k) - a.Units
                        End If
                    Next a
                    k = k + 1
                End If
            Next res
        Next t

        ' === Schedule candidate tasks ===
        For Each candidate In candidateTasks
            overallocatedCount = 0
            m = 1

            For Each res In Project.Resources
                If Not res Is Nothing Then
                    totalUnits = 0
                    For Each a In candidate.Assignments
                        If a.Resource.ID = res.ID Then
                            totalUnits = totalUnits + a.Units
                        End If
                    Next a
                    If totalUnits > availableUnits(m) Then
                        overallocatedCount = overallocatedCount + 1
                    End If
                    m = m + 1
                End If
            Next res

            If overallocatedCount >= 1 Then
                newDate = DateAdd("d", 1, currentDate)
                candidate.Start = newDate
            Else
                k = 1
                For Each res In Project.Resources
                    If Not res Is Nothing Then
                        For Each a In candidate.Assignments
                            If a.Resource.ID = res.ID Then
                                availableUnits(k) = availableUnits(k) - a.Units
                            End If
                        Next a
                        k = k + 1
                    End If
                Next res
            End If
        Next candidate

        ' === Reset collections ===
        Set candidateTasks = New Collection
        Set ongoingTasks = New Collection

        ' === Update date and status bar ===
        currentDate = currentDate + 1
        Application.StatusBar = "Iterations: " & currentDate - startDate & _
                                " | % Processed Tasks = " & (ongoingTasks.Count + completedTasks.Count) & "/" & Project.Tasks.Count & " = " & _
                                Format((ongoingTasks.Count + completedTasks.Count) / Project.Tasks.Count, "0.00%")

    Wend

    ' === Finalization ===
    macroEndTime = Timer
    macroElapsedTime = macroEndTime - macroStartTime

    Application.StatusBar = "Iterations: " & currentDate - startDate & " | % Processed Tasks = 100%" & " | Rescheduling Completed in " & Format(macroElapsedTime, "0.00") & " seconds"
     
      
End Sub



Function DetectOverallocation() As Boolean
    Dim ProjectResource As Resource

    For Each ProjectResource In ActiveProject.Resources
        If Not ProjectResource Is Nothing Then
            If ProjectResource.Overallocated Then
                DetectOverallocation = True
                Exit Function
            End If
        End If
    Next ProjectResource

    DetectOverallocation = False
    
End Function

