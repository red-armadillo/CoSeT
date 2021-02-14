    Option Explicit
    
    Const SYSTEM_PARAMETERS_SHEET As String = "System Parameters"
    Const SP_MAX_NUM_OF_CRITERIA_RANGE As String = "C2"
    Const SP_MAX_PROJECTS_CELL As String = "C3"
    Const SP_MAX_NUMBER_OF_MARKERS_CELL As String = "C4"
    Const SP_MAX_NUMBER_OF_ASSIGNMENTS_PER_MARKER As String = "C5"
    Const SP_MAX_NUMBER_OF_MARKERS_PER_PROJ As String = "C6" 'where to max # markers per project supported
    Const SP_MAX_KEYWORDS_CELL As String = "C7"
    Const SP_LOCKED_SHEET_PWD_CELL As String = "C8"         ' where to find the sheet password on the parameters sheet.
    Const SP_PROJECT_EXPERTISE_FILE_PATTERN = "C9"          ' used for loading expertise info about the markers
    Const SP_KEYWORD_EXPERTISE_FILE_PATTERN = "C10"          ' used for loading expertise info about the markers
    Const SP_MARKER_SCORING_FILE_PATTERN As String = "C11"   ' for loading the scores from markers
    Const SP_MARKS_WITH_COMMENTS_FILE_PATTERN As String = "C12"
    Const SP_SAME_ORGANIZATION_TEXT_CELL As String = "C13"
    Const SP_SIMULATE_MARKER_RESPONSES_CELL As String = "C14"
    
    Const COMPETITION_PARAMETERS_SHEET As String = "Competition Parameters"
    Const CP_TARGET_MARKERS_PER_PROJ As String = "C3"   'how many markers per project are desired
    Const CP_TARGET_ASSIGNMENTS_PER_MARKER = "C4"       'how many many projects assigned to a marker are desired
    Const CP_NUM_KEYWORDS_CELL As String = "C5"
    Const CP_COMPETITION_ROOT_FOLDER As String = "C10"
    Const CP_EXPERTISE_BY_PROJECT_REQUESTED_FOLDER_CELL As String = "C11"
    Const CP_EXPERTISE_BY_PROJECT_RECEIVED_FOLDER_CELL As String = "C12"
    Const CP_EXPERTISE_BY_KEYWORD_REQUESTED_FOLDER_CELL As String = "C13"
    Const CP_EXPERTISE_BY_KEYWORD_RECEIVED_FOLDER_CELL As String = "C14"
    Const CP_SCORES_REQUESTED_FOLDER_CELL As String = "C15"
    Const CP_SCORES_RECEIVED_CELL As String = "C16"
    Const CP_COMMENTS_FOLDER_CELL As String = "C17"
    Const CP_USE_ORG_DISAMBIGUATION_CELL As String = "C18"
    Const CP_USE_EMAIL_DISAMBIGUATION_CELL As String = "C19"
    Const CP_USE_NORMALIZED_SCORING_CELL = "C20"
    Const CP_GATHER_COMMENTS_CELL = "C21"
    Const CP_COMMENT_OUTPUT_FORMAT As String = "C22"
    Const CP_MAX_FIRST_READER_ASSIGNMENTS_CELL As String = "K15"
    
    Const CRITERIA_SHEET As String = "Criteria"
    Const C_NUMBER_OF_CRITERIA_CELL As String = "G1"
    Const C_FIRST_CRITERIA_MINVALUE_RN As Long = 3
    Const C_FIRST_CRITERIA_MINVALUE_CN As Long = 3
    
    Const PROJECTS_SHEET As String = "Projects"
    Const PS_NUMBER_OF_PROJECTS_CELL As String = "L1"
    Const PS_PROJECT_NAME_COLUMN As Long = 2
    Const PS_CONTACT_NAME_COLUMN As Long = 3
    Const PS_ORG_COLUMN As Long = 4                 ' organization of the submitters
    Const PS_CONTACT_EMAIL_COLUMN As Long = 5
    Const PS_MENTOR_ID_COLUMN As Long = 7
    Const PS_FIRST_DATA_ROW As Long = 3
    
    Const MARKERS_SHEET As String = "Markers"      ' sheet containg the marker names and their associated marker number
    Const M_NUM_MARKERS_CELL As String = "G1"      ' cell containing a count of all the people who registered as markers
    Const M_NUMBER_AND_NAME_COLUMNS As String = "A:B"
    Const M_ORG_COLUMN As Long = 3
    Const M_EMAIL_COLUMN As Long = 4
    Const M_NUM_TEAMS_MENTORED_COLUMN As Long = 5
    Const M_FIRST_DATA_ROW As Long = 2
    
    Const KEYWORDS_SHEET As String = "Keywords"
    Const KEYWORD_STRING As String = "Keyword"
    Const KW_NUM_KEYWORDS_CELL As String = "G2"
    Const KW_KEYWORDS_COL As Long = 3
    Const KW_WEIGHTS_COL As Long = 4
    Const KW_WEIGHTS_ROW As Long = 3
    
    Const PROJECT_KEYWORDS_SHEET As String = "Project Keywords"
    Const PK_FIRST_PROJECT_DATA_COL As Long = 3
    Const PK_FIRST_PROJECT_DATA_ROW As Long = 4
    Const PK_FIRST_NORMALIZED_DATA_COL As Long = 36
    
    Const MARKER_EXPERTISE_SHEET As String = "Marker Expertise"
    Const ME_FIRST_MARKER_DATA_COL As Long = 3
    Const ME_FIRST_MARKER_DATA_ROW As Long = 4
    Const ME_FIRST_NORMALIZED_DATA_COL As Long = 36

    Const PROJECT_X_MARKER_SHEET As String = "Project X Marker Table"
    Const PXM_FIRST_DATA_ROW As Long = 4
    Const PXM_FIRST_PXM_COL As Long = 4
    Const PXM_MARKER_NUM_ROW As Long = 1
    
    Const EXPERTISE_BY_PROJECTS_INSTRUCTIONS_SHEET As String = "Expertise by Projects - Instr."
    
    Const EXPERTISE_CROSSWALK_SHEET As String = "Expertise Crosswalk"
    Const EC_DATA_FIRST_MARKER_COLUMN As Long = 31   ' aka "AE"
    Const EC_DATA_FIRST_MARKER_ROW As Long = 7
    Const EC_XLMH_CONFIDENCE_PER_PROJECT_COL As Long = 5        ' 4 col. with COIs, low, medium, high about a project
    Const EC_ASSIGNMENT_CONFIDENCE_FIRST_COLUMN As Long = 10    ' col. with project confidence for marker assigned
    Const EC_ASSIGNMENTS_FIRST_COLUMN As Long = 20              ' col. with marker numbers assigned to a project
    Const EC_XLMH_MARKER_TABLE_FIRST_ROW As Long = 2            ' 4 rows with COIs, low, medium, high about a marker
        
    Const MASTER_ASSIGNMENTS_SHEET As String = "Assignments Master"
    Const MAS_FIRST_ASSIGNMENT_COLUMN As Long = 6               ' column "F" columns that span the marking assignments
    Const MAS_FIRST_ASSIGNMENT_ROW As Long = 3
    Const MAS_FLAG_COLUMN As String = "AJ"
    Const MASTER_SCORESHEET As String = "Scoresheet Master"
    Const MSS_PROJECT_COLUMN As Long = 1
    Const MSS_FIRST_PROJECT_ROW As Long = 6
    Const MSS_FIRST_SCORE_COL As Long = 9
    Const MSS_TOTAL_SCORES_COLUMN As Long = 6       ' column with the total scores for a project
    Const MSS_LAST_COL As Long = 127                 ' last column of the master storing sheet
    Const MSS_MARKER_NUMBER_ROW As Long = 2         ' for "Marker #N (Normalized/Raw) Criteria Scores"
    
    Const SHARED_SCORESHEET As String = "Shared Scoresheet"
    Const SS_FIRST_DATA_ROW As Long = 5
    Const SS_PROJECT_NUM_COLUMN As Long = 3
    Const SS_MARKER_NUM_COLUMN As Long = 1
    Const SS_FIRST_SORT_COLUMN As Long = 1
    Const SS_LAST_SORT_COLUMN As Long = 17
    Const SS_FIRST_RAW_COLUMN As Long = 6
    Const SS_FIRST_NORMAL_COLUMN As Long = 30
    Const SS_FIRST_FINAL_COLUMN As Long = 44
    Const SS_FINAL_PROJ_COLUMN As Long = 42
    Const SS_FINAL_TOTAL_SCORES_COLUMN As Long = 55
    Const SHARED_SCORESHEET_TEMPLATE As String = "Shared Scoresheet - template"
    
    Const EBPI_FOR_MORE_INFO_ROW As Long = 6
    
    Const EXPERTISE_BY_KEYWORDS_INSTRUCTIONS_SHEET As String = "Expertise by Keywords - Instr."
    Const EBKI_FOR_MORE_INFO_ROW As Long = 3
    
    Const MARKER_PROJECT_EXPERTISE_TEMPLATE As String = "Marker Project - template"
    Const MPET_FIRST_DATA_ROW As Long = 2
    Const MPET_COI_COLUMN As Long = 6
    Const MPET_EXPERTISE_COLUMN As Long = 7
    Const MPET_MARKER_INFO_COLUMN As String = "K"
    
    Const MARKER_KEYWORD_EXPERTISE_TEMPLATE As String = "Marker Keyword - template"
    Const MKET_EXPERTISE_COLUMN As String = "C"
    Const MKET_FIRST_DATA_ROW As Long = 2
    Const MKET_MARKER_NAME_CELL As String = "F1"
    Const MKET_MARKER_NUM_CELL As String = "F2"
    Const MKET_COI_SHEET_NAME As String = "Conflicts of Interest"
    
    Const MARKER_SCORING_TEMPLATE_SHEET As String = "Marker Scoresheet - template"
    Const MST_FIRST_SCORING_ROW As Long = 9             ' row of first project score in marker' sheet
    Const MST_FIRST_SCORING_COL As Long = 4
    Const MST_MARKER_NAME_CELL As String = "C1"
    Const MST_MARKER_NUMBER_CELL As String = "Y2"
    Const MST_FIRST_NORMALIZED_SCORE_COLUMN As Long = 15
    Const MST_NAME_COL As String = "B"
    Const MST_PROJECT_NUM_COL As String = "A"
    Const MST_READER_NUM_COL As String = "C"
    Const MST_PROJECT_COUNT_CELL As String = "N60"      ' location of # of projects on a score sheet
    Const MST_TARGET_SCORING_FRACTION_CELL As String = "O61"
    Const MST_EXPECTED_NUMBER_OF_SCORES As String = "E8"
    Const MST_CRITERIA_NAMES_AND_SCORE_RANGES = "D5:M7"
    
    Const SCORES_AND_COMMENTS_INSTRUCTIONS_TEMPLATE As String = "Instructions to Markers"
    Const SCI_INSTRUCTION_SHEET_NAME As String = "Instructions"
    Const SCI_COMPETITION_NAME_CELL As String = "A1"
    Const SCI_FOR_MORE_INFO_CELL As String = "A4"
    Const SCI_PROJECT_COUNT_ROW As Long = 6
    
    Const SCORES_AND_COMMENTS_TEMPLATE_SHEET As String = "Scores and Comments - template"
    Const SCT_COMPETITION_NAME_CELL = "B1"
    Const SCT_MARKER_NAME_CELL As String = "B2"
    Const SCT_MARKER_NUM_CELL As String = "B3"
    Const SCT_PROJECT_NUM_CELL As String = "B4"
    Const SCT_PROJECT_NAME_CELL As String = "B5"
    Const SCT_CRITERIA_ONE_NAME_CELL As String = "D12"
    Const SCT_CRITERIA_ONE_MIN_CELL As String = "B13"
    Const SCT_CRITERIA_ONE_MAX_CELL As String = "D13"
    Const SCT_READER_NUM_CELL As String = "E4"
    Const SCT_FIRST_CRITERIA_SCORE As String = "F13"
    Const SCT_SCORE_CHECK_CELL As String = "G2"
    Const SCT_COI_RESPONSE_CELL As String = "K6"
    Const SCT_CONFIDENCE_LOW_CELL As String = "B8"
    Const SCT_CONFIDENCE_MEDIUM_CELL As String = "E8"
    Const SCT_CONFIDENCE_HIGH_CELL As String = "H8"
    Const SCT_SCORE_COLUMN As Long = 6
    Const SCT_GENERAL_COMMENT_CELL As String = "A10"
    Const SCT_ROWS_PER_CRITERIA As Long = 5
    Const SCT_FULL_SHEET_RANGE As String = "A1:K61"
    
    Const PROJECT_COMMENTS_SHEET As String = "Project Comments - template"
    Const PC_PROJECT_NUM_CELL As String = "B2"
    Const PC_PROJECT_NAME_CELL As String = "B3"
    Const PC_GENERAL_COMMENTS_CELL As String = "A5"
    Const PC_FIRST_CRITERIA_COMMENTS_ROW As Long = 9
    Const PC_NUM_ROWS_PER_COMMENT As Long = 4
    
    Const ONE_THIRD As Double = 1 / 3
    Const TWO_THIRDS As Double = 2 / 3
    Const MAX_EMAIL_LENGTH As Long = 30
    
' GLOBALS GLOBALS GLOBALS GLOBALS GLOBALS GLOBALS GLOBALS GLOBALS GLOBALS GLOBALS GLOBALS GLOBALS
    Global col_ltrs() As String
    Global a2z() As String
    
    Global main_workbook As Variant
    
    Global max_criteria As Long
    Global num_criteria As Long
    Global max_projects As Long
    Global num_projects As Long
    Global max_markers As Long
    Global num_markers As Long                  ' number of people from the Markers sheet (calculated)
    Global max_markers_per_proj As Long         ' max # of markers per project supported by this set of tools
    Global target_markers_per_proj As Long      ' # of markers as specified in the competition parameters sheet
    Global max_ass_per_marker As Long
    Global target_ass_per_marker As Long
    Global max_first_reader_assignments As Long
    Global n_per_assignment_col() As Long       ' number of readers assigned at this level
    Global retain_worksheets As Boolean
    Global num_keywords As Long
    Global max_keywords As Long
    Global normalize_scoring As Boolean
    Global project_expertise_file_pattern As String
    Global keyword_expertise_file_pattern As String
    Global expertise_by_project_ending As String
    Global expertise_by_keyword_ending As String
    Global ss_marks_file_pattern As String            'e.g., * marks.xlsx
    Global ss_marks_comments_file_pattern As String
    Global same_organization_text As String
    Global use_email_disambiguation As Boolean      ' for disambiguating personalized files
    Global use_org_disambiguation As Boolean
    Global gather_comments As Boolean
    Global output_comments_format As String
    
    Global root_folder As String
    Global expertise_by_project_requested_folder As String
    Global expertise_by_project_received_folder As String
    Global expertise_by_keyword_requested_folder As String
    Global expertise_by_keyword_received_folder As String
    Global scores_received_folder As String
    Global scores_requested_folder As String
    Global comments_folder As String
    
    Global scores_ending As String          ' for sheets with only scores
    Global scores_with_comments_ending As String        ' for sheets with comments and scores
    
    ' Arrays for assigning markers
    Global mc_array() As Variant            'mc_ for Marker Confidence (gets updated by assignments)
    Global mc_as_loaded() As Variant        'as read in from the expertise crosswalk
    Global pn_array() As Variant            ' pn for project number (project numbers can be in random order)
    Global mn_array() As Variant            ' mn for marker number (marker numbers can be in random order)
    Global coa_array() As Variant           ' confidence of the assigned marker array
        ' dimension 1 = rows, one project's assigned markers' COA
        ' dimension 2 has N items, one for each possible marker's confidence
    Global mentor_column() As Variant       ' column of COI text for expertise sheet
    Global competition_COIs() As Variant    ' array of COIs for all markers on all projects
    Global ss_marker_col() As Variant       ' the marker# column in the raw table
    Global ss_project_col() As Variant      ' the project# column in the raw table
    
    Global xlmh_per_marker() As Variant     'number of eXcluded, Low, Medium & High selections for a marker
    Global xlmh_per_project() As Variant    'number of eXcluded, Low, Medium & High markers for a project
    Global assignments() As Variant         ' markers/readers assigned to the projects
        ' dimension 1 = rows, one per project
        ' dimension 2 has N items, with the number of each marker assigned
    Global n_assigned2project() As Long        ' count on the number of markers assigned to a project
    Global n_assigned2marker() As Long      ' count on the number of projects assigned to a marker
    Global marker_orgs() As Variant         ' array of organization affiliation provided for markers
    Global marker_emails() As Variant       ' array of emails provided for markers
    Global comments() As Variant            ' array of comments from each marker for each criteria on each project
                                            ' Projects x Criteria
    Global general_comments() As Variant    ' array of the general comments provided by markers (by projects)
    
    Global assignment_failed_for_this_proj() As Boolean
    
    ' these 3 arrays are used to store scores, moved to global to see if it avoids crashes.
    Global scores() As Variant, project_nums() As Variant, reader_nums() As Variant
    
    Global messages() As String             ' for buffering messages
    Global num_messages As Long
    Global buffer_messages As Boolean
    
    Global simulate_marker_responses As Boolean
    
    
    Global globals_defined As Boolean
    
' Debug
    Dim ass_this_col() As Long
' end debug

' bundle the middle steps together
Sub Expertise2MarkingSheets()

    InitMessages
    
    If LoadMarkerExpertiseIntoCrosswalk = False Then
        ReportMessages
        Exit Sub
    End If
    
    
    If AssignMarkersBasedOnConfidence = False Then
        ReportMessages
        Exit Sub
    End If
    
    If CreateAllMarkingSheets = False Then
        ReportMessages
        Exit Sub
    End If
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

' THIS APPEARS TO MAKE THE EXCEL SESSION HANG
'    If ExportCompetitionWorkbook = False Then
'        ReportMessages
'        Exit Sub
'    End If
    
    ReportMessages
    
End Sub

Public Function KeywordTablesToScoresheets() As Boolean

    
    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    If (CreatePXMFromProjectRelevanceAndMarkerExpertise = False) Then
        Exit Function
    End If
    
    If AssignMarkersBasedOnConfidence = False Then
        ReportMessages
        Exit Function
    End If
    
    If CreateAllMarkingSheets = False Then
        ReportMessages
        Exit Function
    End If
    
End Function

Sub UnsortMasterScoresheet()
'
' UnsortMasterScoresheet Macro
'
    DefineGlobals
    Dim start_address As String
    start_address = ActiveCell.Address

    Const sort_key_col As String = "A"
    Dim sort_range As String
    sort_range = sort_key_col & MSS_FIRST_PROJECT_ROW & ":" & _
                  c2l(MSS_LAST_COL + 2) & (MSS_FIRST_PROJECT_ROW + max_projects - 1)
    Range(FirstCell(sort_range)).Select
    Range(FirstCell(sort_range)).Activate
    Dim first_row As Long, sort_key As String
    first_row = ActiveCell.row
    sort_key = sort_key_col & first_row & ":" & sort_key_col & (first_row + max_projects - 1)
    ActiveWorkbook.Worksheets(MASTER_SCORESHEET).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(MASTER_SCORESHEET).Sort.SortFields.Add2 Key:=Range(sort_key), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(MASTER_SCORESHEET).Sort
        .SetRange Range(sort_range)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' go back to where we started
    Range(start_address).Select
End Sub

    

Public Sub MakeProjectExpertiseSheets()
    
    WriteAllExpertiseSheets expertise_by_project_requested_folder, expertise_by_project_ending, ""

End Sub

Public Sub MakeKeywordExpertiseSheets()
    
    WriteAllExpertiseSheets expertise_by_keyword_requested_folder, expertise_by_keyword_ending, KEYWORD_STRING

End Sub

Public Function WriteAllExpertiseSheets(expertise_type As String, expertise_type_ending As String, _
                                        keyword_or_project As String) As Boolean
    'Turn off events and screen flickering.
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim start_range As String
    start_range = ActiveCell.Address
    
    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    ' position the cursor on the appropriate sheet, at at the top of the list of marker #'s
    Dim starting_sheet_name As String, n_expertise_books As Long
    starting_sheet_name = ActiveSheet.Name
    Dim cell_name As String
    Sheets(MARKERS_SHEET).Activate
    
    ' as long as we have a marker number and marker name create a marker expertise sheet
    Dim marker_name As String, marker_number As Long
    Dim i As Long
    Range("A1").Select
    
    Dim folder_for_expertises As String
'    ChDir root_folder
    folder_for_expertises = SelectFolder("Select folder to store the blank" & expertise_type_ending & " sheets", _
                root_folder & expertise_type)
    If Len(folder_for_expertises) = 0 Then
        Exit Function
    End If

    ' look through the rows on the markers sheet and create an expertise workbook for each person
    For i = 1 To num_markers
        ThisWorkbook.Activate
        Sheets(MARKERS_SHEET).Activate
        ChangeActiveCell 1, 0             ' move right to check the marker's name
        marker_number = ActiveCell.Value
        ChangeActiveCell 0, 1             ' move right to check the marker's name
        marker_name = ActiveCell.Value
        ChangeActiveCell 0, -1            ' move right to check the marker's name
        If (marker_number <> i) Or (Len(marker_name) = 0) Then
           AddMessage "[writeAllExpertiseSheets] expected marker number and name, got [" & marker_number _
                    & "] and {" & marker_name & "} - exiting."
            Exit Function
        End If
        If keyword_or_project = KEYWORD_STRING Then
            WriteOneKeywordExpertiseBook marker_number, marker_name, folder_for_expertises
        Else
            WriteOneProjectExpertiseBook marker_number, marker_name, folder_for_expertises
        End If
        n_expertise_books = n_expertise_books + 1
    Next i
    
    Sheets(starting_sheet_name).Activate
    Range(start_range).Select
    Range(start_range).Activate
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    AddMessage "Created " & n_expertise_books & " sheets for markers to provide their confidence about marking projects."

End Function

Public Function WriteOneProjectExpertiseBook(marker_num As Long, marker_name As String, folder As String) As String

    ' create the workbook that has the expertise sheet and associated instructions for each expert
    Dim new_sheet As String
        
    ' move the new sheet out into its own workbook
    Sheets(Array(EXPERTISE_BY_PROJECTS_INSTRUCTIONS_SHEET, MARKER_PROJECT_EXPERTISE_TEMPLATE)).Select
    Sheets(MARKER_PROJECT_EXPERTISE_TEMPLATE).Activate
    Sheets(Array(EXPERTISE_BY_PROJECTS_INSTRUCTIONS_SHEET, MARKER_PROJECT_EXPERTISE_TEMPLATE)).Copy
    
    ' remove the external link in the instruction sheet
    Sheets(EXPERTISE_BY_PROJECTS_INSTRUCTIONS_SHEET).Activate
    ConvertRangeToText "A1"
    ConvertRangeToText "A" & EBPI_FOR_MORE_INFO_ROW
    
    Sheets(MARKER_PROJECT_EXPERTISE_TEMPLATE).Select
    ' store the marker number and name on this sheet
    Range(MPET_MARKER_INFO_COLUMN & 1).Value = marker_num
    Range(MPET_MARKER_INFO_COLUMN & 2).Value = marker_name
    ' remove external links in the expertise template
    ConvertRangeToText "A1:E1"          ' column headers
    ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "A"
    ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "B"
    ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "C"
    ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "D"
    ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "E"
    ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "F"
    
    ' rename the expertise sheet to the marker's number and name
    ActiveSheet.Name = GoodTabName(marker_num & " " & marker_name)
    new_sheet = ActiveSheet.Name
    
    If simulate_marker_responses Then
        ' fill the expertise column of the sheet with made up high-medium-low
        ' so they can be used for assignment testing (but respect the COI info)
        Dim COI_range As String, COI_column() As Variant
        COI_range = c2l(MPET_COI_COLUMN) & MPET_FIRST_DATA_ROW & ":" & _
                    c2l(MPET_COI_COLUMN) & (MPET_FIRST_DATA_ROW + num_projects - 1)
        COI_column = Range(COI_range)
        Range(c2l(MPET_EXPERTISE_COLUMN) & MPET_FIRST_DATA_ROW).Select
        WriteRandomExpertiseRatings num_projects, True, COI_column
    End If
    
    ' lock the sheet
    LockUserProjectExpertiseSheet ActiveSheet.Name, MPET_EXPERTISE_COLUMN
    
    ' if requested, insert the marker's organization and email into the filename
    Dim file_stub As String
    file_stub = DisambiguateFilename(new_sheet, marker_num)
    ' write this workbook to the current directory and close it.
    WriteOneProjectExpertiseBook = folder & "\" & file_stub & expertise_by_project_ending & ".xlsx"
    ActiveWorkbook.SaveAs filename:=WriteOneProjectExpertiseBook, _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
    
End Function

Public Function WriteOneKeywordExpertiseBook(marker_num As Long, marker_name As String, folder As String) As String

    ' create the workbook containing the expertise sheet and associated instructions
    Dim keyword_sheet As String, COI_sheet As String
        
    ' move the required sheets out into their own workbook
    Sheets(Array(EXPERTISE_BY_KEYWORDS_INSTRUCTIONS_SHEET, MARKER_KEYWORD_EXPERTISE_TEMPLATE, _
                MARKER_PROJECT_EXPERTISE_TEMPLATE)).Select
    Sheets(MARKER_KEYWORD_EXPERTISE_TEMPLATE).Activate
    Sheets(Array(EXPERTISE_BY_KEYWORDS_INSTRUCTIONS_SHEET, MARKER_KEYWORD_EXPERTISE_TEMPLATE, _
                MARKER_PROJECT_EXPERTISE_TEMPLATE)).Copy
    
    ' remove the external link in the instruction sheet
    Sheets(EXPERTISE_BY_KEYWORDS_INSTRUCTIONS_SHEET).Activate
    ConvertRangeToText "A1"
    ConvertRangeToText "A" & EBKI_FOR_MORE_INFO_ROW
    
    ' remove external links from the keyword rating sheet
    Sheets(MARKER_KEYWORD_EXPERTISE_TEMPLATE).Select
    Sheets(MARKER_KEYWORD_EXPERTISE_TEMPLATE).Activate
    ConvertCellsDownFromFormula2Text MKET_FIRST_DATA_ROW, "A"
    ConvertCellsDownFromFormula2Text MKET_FIRST_DATA_ROW, "B"
    
    ' store the marker number and name on this sheet
    Range(MKET_MARKER_NUM_CELL).Value = marker_num
    Range(MKET_MARKER_NAME_CELL).Value = marker_name
    
'   update what has become the COI sheet
    Sheets(MARKER_PROJECT_EXPERTISE_TEMPLATE).Select
    Sheets(MARKER_PROJECT_EXPERTISE_TEMPLATE).Activate
    Range(MPET_MARKER_INFO_COLUMN & 1).Value = marker_num
    Range(MPET_MARKER_INFO_COLUMN & 2).Value = marker_name
    ConvertRangeToText "A1:E1"
    ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "A"
    ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "B"
    ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "C"
    ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "D"
    ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "E"
    ConvertCellsDownFromFormula2Text MPET_FIRST_DATA_ROW, "F"
    
    ' remove the expertise column, and the error checking column from the project expertise sheet
    ' with that, the sheet becomes a COI request sheet.
    Dim expertise_col As Long
    expertise_col = FindHeaderColumn(1, "Expertise:", False)
    If expertise_col = 0 Then
        MsgBox "WriteOneKeywordExpertiseBook: unable to find Expertise: column", vbCritical
        Exit Function
    End If
    Columns(c2l(expertise_col) & ":" & c2l(expertise_col + 1)).Select
    Selection.Delete Shift:=xlToLeft
    ActiveSheet.Name = "Conflicts of Interest"
    COI_sheet = ActiveSheet.Name
    ' Also personalize this sheet with the marker number and name
    
    ' name the expertise sheet to the marker's number and name
    Sheets(MARKER_KEYWORD_EXPERTISE_TEMPLATE).Select
    ActiveSheet.Name = GoodTabName(marker_num & " " & marker_name)
    keyword_sheet = ActiveSheet.Name
    
    If simulate_marker_responses Then
        ' fill the expertise column of the sheet with made up high-medium-low
        ' so they can be used for assignment testing (but respect the COI info)
        Dim COI_column() As Variant
        Range(MKET_EXPERTISE_COLUMN & MKET_FIRST_DATA_ROW).Select
        WriteRandomExpertiseRatings num_keywords, False, COI_column
    End If
            
    ' write this workbook to the current directory and close it.
    LockUserKeywordExpertAndCOISheets keyword_sheet, COI_sheet, expertise_col - 2
    Sheets(1).Select    ' make the instructions sheet visible
    Sheets(1).Activate
    
    ' if requested, insert the marker's organization and email into the filename
    Dim file_stub As String
    file_stub = DisambiguateFilename(keyword_sheet, marker_num)
    ' now save the file to disk
    WriteOneKeywordExpertiseBook = folder & "\" & file_stub & expertise_by_keyword_ending & ".xlsx"
    ActiveWorkbook.SaveAs filename:=WriteOneKeywordExpertiseBook, _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
    
End Function

Public Function DisambiguateFilename(filestub As String, marker_num As Long) As String
    

    Dim cell_range As String
    If use_org_disambiguation Then
        Dim org As String, marker_org As String
        cell_range = c2l(M_ORG_COLUMN) & (M_FIRST_DATA_ROW + marker_num - 1)
        marker_org = ThisWorkbook.Sheets(MARKERS_SHEET).Range(cell_range).Value
        'read the org name for this expert
        org = Email2Text(marker_org, MAX_EMAIL_LENGTH)
        ' if the org is non-null tack it on
        If Len(org) > 0 Then
            filestub = filestub & " " & org
        End If
    End If
    If use_email_disambiguation Then
        Dim email As String, marker_email As String
        cell_range = c2l(M_EMAIL_COLUMN) & (M_FIRST_DATA_ROW + marker_num - 1)
        marker_email = ThisWorkbook.Sheets(MARKERS_SHEET).Range(cell_range).Value
        'read the email address for this expert
        email = Email2Text(marker_email, MAX_EMAIL_LENGTH)
        ' if the email is non-null tack it on
        If Len(email) > 0 Then
            filestub = filestub & " " & email
        End If
    End If
    DisambiguateFilename = filestub
    
End Function


Public Function LockUserProjectExpertiseSheet(project_expertise_sheet As String, lc As Long) As Boolean
    'lock most of the sheet for recording COIs
    Sheets(project_expertise_sheet).Select
    Dim lock_range As String, free_range As String
    lock_range = "A:" & c2l(lc - 2) & "," & c2l(lc + 1) & ":" & c2l(lc + 1) & "," & _
                c2l(lc + 3) & "1:" & c2l(lc + 4) & "2," & c2l(lc) & "1"
    Range(lock_range).Select
    Range("A1").Activate
    Selection.Locked = True
    Selection.FormulaHidden = False
    
    ' make sure the areas for data entry are not locked.
    free_range = c2l(lc - 1) & MPET_FIRST_DATA_ROW & ":" & c2l(lc) & (MPET_FIRST_DATA_ROW - 1 + num_projects)
    Range(free_range).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Locked = False
    Selection.FormulaHidden = False
    
    ' lock the sheet
    Dim pwd As String
    pwd = Workbooks(main_workbook).Sheets(SYSTEM_PARAMETERS_SHEET).Range(SP_LOCKED_SHEET_PWD_CELL).Value
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=pwd

    LockUserProjectExpertiseSheet = True
End Function
Public Function LockUserKeywordExpertAndCOISheets(keyword_sheet As String, COI_sheet As String, _
                     lc As Long) As Boolean
'
    'lock most of the sheet for recording COIs
    Sheets(COI_sheet).Select
    Dim lock_range As String, free_range As String
    lock_range = "A:" & c2l(lc) & "," & c2l(lc + 1) & "1," & c2l(lc + 3) & "1:" & c2l(lc + 4) & "2"
    Range(lock_range).Select
    Range("A1").Activate
    Selection.Locked = True
    Selection.FormulaHidden = False
    
    ' make sure the areas for data entry are not locked.
    free_range = c2l(lc + 1) & MKET_FIRST_DATA_ROW & ":" & c2l(lc + 1) & (MKET_FIRST_DATA_ROW - 1 + num_projects)
    Range(free_range).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Locked = False
    Selection.FormulaHidden = False
    
    ' lock the sheet
    Dim pwd As String
    pwd = Workbooks(main_workbook).Sheets(SYSTEM_PARAMETERS_SHEET).Range(SP_LOCKED_SHEET_PWD_CELL).Value
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=pwd
    
    'now protect most of the Keyword sheet
    Sheets(keyword_sheet).Select
    lock_range = "A:B,E1:F2"
    Range(lock_range).Select
    Range("C2").Activate
    Selection.Locked = True
    Selection.FormulaHidden = False
    ' make sure the areas for data entry are not locked.
    free_range = MKET_EXPERTISE_COLUMN & MKET_FIRST_DATA_ROW & ":" & _
          MKET_EXPERTISE_COLUMN & MKET_FIRST_DATA_ROW - 1 + max_keywords
    Range(free_range).Select
    Selection.Locked = False
    Selection.FormulaHidden = False
    
    ' lock the sheet
    pwd = Workbooks(main_workbook).Sheets(SYSTEM_PARAMETERS_SHEET).Range(SP_LOCKED_SHEET_PWD_CELL).Value
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=pwd

End Function

Public Function random_expertise_rating_LMH() As String
    ' create a random rating letter grade: L, M or H, with the bias indicated by the case statement below.
    Const upperbound As Long = 100
    Const lowerbound As Long = 1
    Dim rating_letters(1 To 3) As String, rating As Long
    rating_letters(1) = "L"
    rating_letters(2) = "M"
    rating_letters(3) = "H"
    rating = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
    Select Case rating  '80 L, 17 M, 3 H
    Case 1 To 80
        random_expertise_rating_LMH = rating_letters(1)
    Case 81 To 97
        random_expertise_rating_LMH = rating_letters(2)
    Case 98 To 100
        random_expertise_rating_LMH = rating_letters(3)
    Case Else
        MsgBox "Error in case statement", vbCritical
    End Select
    
End Function

Function WriteRandomExpertiseRatings(num_ratings As Long, check_COI As Boolean, COI_column() As Variant) As Boolean
    ' write a column of random expertise ratings starting at the current location
    Dim i As Long, rating As Long
    Dim ratings_column() As String
    ReDim ratings_column(1 To num_ratings)
    ' create the random numbers
    For i = 1 To num_ratings
        ' bias the confidence ratings to low, then medium, then high
        If check_COI Then
            If Len(COI_column(i, 1)) = 0 Then
                ratings_column(i) = random_expertise_rating_LMH()
            Else
                ' this marker is in conflict for this project, flag it.
                ratings_column(i) = "X"
            End If
        Else
            ratings_column(i) = random_expertise_rating_LMH()
        End If
    Next i
    
    ' now write the array to the expertise column of the expertise sheet
    Dim Destination As Range
    Set Destination = Range(ActiveCell.Address)
    Set Destination = Destination.Resize(UBound(ratings_column), 1)
    Destination.Value = Application.Transpose(ratings_column)
    
    WriteRandomExpertiseRatings = True
End Function

Public Function DuplicateTemplateSheet(template_name As String) As String ' returns the name of the new sheet
' make a duplicate of a given sheet and make it the active sheet
' new sheet's name (as created by Excel) is returned
    On Error GoTo duplicating_sheet_error
    
    Sheets(template_name).Copy Before:=Sheets(Sheets.Count)
    DuplicateTemplateSheet = ActiveSheet.Name
    Exit Function
duplicating_sheet_error:
    MsgBox "[DuplicateTemplateSheet] Error duplicating sheet {" & template_name & "}, check sheet exists", vbCritical
    DuplicateTemplateSheet = ""
    On Error GoTo 0
    Exit Function
End Function

Public Function GoodTabName(name_in As String) As String
    Dim i As Long, j As Long, name_out As String, one_char As String
    Const MAX_TAB_NAME_LENGTH As Long = 30
    
    For i = 1 To Len(name_in)
        one_char = Mid(name_in, i, 1)
        If ((one_char >= "a") And (one_char <= "z")) Or _
        ((one_char >= "A") And (one_char <= "Z")) Or _
        ((one_char >= "0") And (one_char <= "9")) Or _
        (one_char = "_") Or (one_char = "-") Or (one_char = " ") Then
            j = j + 1
            name_out = name_out & one_char
        End If
    Next i
    If j = 0 Then name_out = "BAD TAB NAME"
    If Len(name_out) > MAX_TAB_NAME_LENGTH Then
        name_out = Left(name_out, MAX_TAB_NAME_LENGTH)
    End If
    GoodTabName = name_out
End Function

Sub test_createallmarkingsheets()
    CreateAllMarkingSheets
End Sub
Public Function CreateAllMarkingSheets() As Boolean
    ' create the sheets for markers to enter their scoring of projects

    'make sure we are starting in the right sheet
    Dim starting_sheet As String
    starting_sheet = ActiveSheet.Name
    Sheets(MASTER_ASSIGNMENTS_SHEET).Activate
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If

    Dim num_scoring_files_created As Long
    Dim new_sheet As String
    Dim mn As Long
    Dim cell_name As String
    Dim project_num As Long, project_name As String     ' as copied from the marker sheet
    
    ' check the number of markers
    If num_markers < 1 Then
        MsgBox "[CreateAllMarkingSheets] expected positive number of markers, found " & num_markers, vbOKOnly
        Exit Function
    End If
    
    ' move to the column with marker assignments
    Dim first_assignment_row As Long
    Dim row_num As Long, col_num As Long, marker_table_row As Long
    Dim find_range As String, marker_name As String
    Dim assignment_col As Long  '
    Dim marker_workbook_started As Boolean
    
    Range(c2l(MAS_FIRST_ASSIGNMENT_COLUMN + 1) & 1).Select
    find_range = c2l(MAS_FIRST_ASSIGNMENT_COLUMN) & ":" & c2l(MAS_FIRST_ASSIGNMENT_COLUMN + target_markers_per_proj - 1)
    
    ' get the folder to store output in
    scores_requested_folder = SelectFolder("Specify where to save the blank scoresheets", _
                            root_folder & scores_requested_folder)
    If Len(scores_requested_folder) = 0 Then
        Exit Function
    End If
    
    ' for each marker make a sheet with their assigned projects
    num_scoring_files_created = 0
    For mn = 1 To num_markers
        marker_workbook_started = False
        new_sheet = ""
        marker_name = GetMarkerName(mn)

        'look through the marker assignment columns for projects they are assigned to, and add them to the sheet
        Dim finding As Boolean, num_assignments As Long, found_cell_label As String
        Dim instructions_sheet As String
        If gather_comments Then
            ' prepare an array with the names of the sheets that will be the marking workbook
            Dim marking_sheets() As String
            ReDim marking_sheets(1 To max_ass_per_marker + 1)
            Dim scores_and_comments_sheet As String
        End If
        
        first_assignment_row = -1
        finding = True
        num_assignments = 0
        found_cell_label = ""
        While finding
            ' find the next assignment for this marker
            Sheets(MASTER_ASSIGNMENTS_SHEET).Activate
            Columns(find_range).Select
            If (Len(found_cell_label) > 0) Then
                Range(found_cell_label).Activate ' move to the last assignment found for this marker
            Else
                ' start at the end of the header of marker numbers (i.e. next cell is first project's first assignment
                Range(c2l(MAS_FIRST_ASSIGNMENT_COLUMN + target_markers_per_proj - 1) & 2).Activate
            End If
            If Selection.Find(What:=mn, after:=ActiveCell, LookIn:=xlFormulas2, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False) Is Nothing Then
                'nothing found, nothing for this marker so exit the loop
                finding = False
            Else
                If marker_workbook_started = False Then     ' first assignment for a marker, start a workbook
                    marker_table_row = MST_FIRST_SCORING_ROW
                    If gather_comments Then
                        ' Eventually we'll make a separate workbook for the submissions this marker is to evaluate
                        ' for now, start with the instructions sheet
                        instructions_sheet = DuplicateTemplateSheet(SCORES_AND_COMMENTS_INSTRUCTIONS_TEMPLATE)
                        ActiveSheet.Name = SCI_INSTRUCTION_SHEET_NAME
                        marking_sheets(1) = ActiveSheet.Name
                        ' remove external references from the sheet
                        ConvertRangeToText SCI_COMPETITION_NAME_CELL
                        ConvertRangeToText SCI_FOR_MORE_INFO_CELL
                        Sheets(MASTER_ASSIGNMENTS_SHEET).Activate
                    Else
                        ' all the scores go in a single scoresheet - create it as copy of the template
                        ' create the new sheet
                        new_sheet = DuplicateTemplateSheet(MARKER_SCORING_TEMPLATE_SHEET)
                        Sheets(MASTER_ASSIGNMENTS_SHEET).Activate
                    End If
                    marker_workbook_started = True
                End If
                ' find out if this assignment is a 'first-reader'
                Selection.Find(What:=mn, after:=ActiveCell, LookIn:=xlFormulas2, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False).Activate
                assignment_col = ActiveCell.Column - MAS_FIRST_ASSIGNMENT_COLUMN + 1
                row_num = ActiveCell.row
                If num_assignments = 0 Then
                    first_assignment_row = row_num
                Else
                    ' if the find has circled back to the start we are done with this marker
                    If row_num <= first_assignment_row Then
                        finding = False
                    End If
                End If
            End If
            
            If finding Then
                num_assignments = num_assignments + 1
                ' we've found an assignment, put the information in the scoresheet for the marker
                col_num = ActiveCell.Column
                found_cell_label = c2l(col_num) & row_num
                project_num = Range(MST_PROJECT_NUM_COL & row_num).Value
                project_name = Range(MST_NAME_COL & row_num).Value
                If gather_comments Then
                    Dim new_sheet_name As String
                    ' add a sheet for this assignment, and populate it with the marker name and number
                    scores_and_comments_sheet = DuplicateTemplateSheet(SCORES_AND_COMMENTS_TEMPLATE_SHEET)
                    ' name it after the project # and name
                    ActiveSheet.Name = GoodTabName(project_num & " " & project_name)
                    marking_sheets(num_assignments + 1) = ActiveSheet.Name
                    Dim table_range As String
                    table_range = SCI_PROJECT_COUNT_ROW + 1 + num_assignments
                    With Sheets(SCI_INSTRUCTION_SHEET_NAME)
                        .Activate
                        .Range("A" & table_range).Value = project_num
                        .Range("B" & table_range).Value = project_name
                        .Range("C" & table_range).Value = marking_sheets(num_assignments + 1)
                        .Range("C" & table_range).Select
                        ' hyperlink to the tab name
                        .Hyperlinks.Add Anchor:=Selection, Address:="", _
                            SubAddress:="'" & marking_sheets(num_assignments + 1) & "'!A1", _
                            TextToDisplay:=marking_sheets(num_assignments + 1)
                    End With
                    Sheets(marking_sheets(num_assignments + 1)).Activate
                    ' insert the data
                    With ActiveSheet
                        .Range(SCT_MARKER_NUM_CELL).Value = mn
                        .Range(SCT_PROJECT_NUM_CELL).Value = project_num
                        .Range(SCT_READER_NUM_CELL).Value = assignment_col
                    End With
                    If simulate_marker_responses Then
                        ' enter random scores and text in the scoresheet
                        If SimulateScoresAndComments = False Then
                            Exit Function
                        End If
                    End If
                    ' make sure the sheet does not have external links
                    If ConvertMarkerCommentSheetFormulaToText = False Then
                        Exit Function
                    End If
                    'lock most of the sheet and hide rows for criteria segments not used
                    If LockAndCompressScoresAndCommentsSheet = False Then
                        Exit Function
                    End If
                Else
                    ' add a scoring row to the single scoresheet
                    Sheets(new_sheet).Activate
                    'put in the team number and project name
                    Range(MST_PROJECT_NUM_COL & marker_table_row).Value = project_num
                    Range(MST_NAME_COL & marker_table_row).Value = project_name
                    Range(MST_READER_NUM_COL & marker_table_row).Value = assignment_col
                    ' increment the row to put the next assignment
                    marker_table_row = marker_table_row + 1
                    ' move down a row to prepare for the next project
                    ChangeActiveCell 1, 0
                End If
            End If
        Wend
        
        ' all done for this marker, set up the outputs for the marker, and export the sheet(s)
        ' if there are assignments for this marker
        If num_assignments > 0 Then
            ' hide the blank rows in the scoresheet
            Dim filestub As String
            If gather_comments Then
                ' we have a multi-sheet workbook (instructions + 1 sheet per submission to evaluate)
                ' move the sheets to a new book, rename it, save and close it
                ReDim Preserve marking_sheets(1 To num_assignments + 1)
                Sheets(marking_sheets(1)).Activate
                Range("B" & SCI_PROJECT_COUNT_ROW).Value = num_assignments
                Sheets(marking_sheets).Select
                Sheets(marking_sheets).Move
                Dim new_workbook As String
                new_workbook = ActiveWorkbook.Name
                filestub = DisambiguateFilename(GoodTabName(mn & " " & marker_name), mn)
                ActiveWorkbook.SaveAs filename:= _
                    scores_requested_folder & "\" & filestub & scores_with_comments_ending & ".xlsx", _
                    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                ActiveWindow.Close
            Else
                If marker_table_row - MST_FIRST_SCORING_ROW + 2 < max_ass_per_marker Then
                    ' there are enough blank rows it is worth doing
                    Sheets(new_sheet).Activate
                    Rows((marker_table_row) & ":" & (MST_FIRST_SCORING_ROW + max_ass_per_marker - 1)).Select
                    Selection.EntireRow.Hidden = True
                End If
                ' put the marker's number and name in the sheet (top row)
                Range(MST_MARKER_NAME_CELL).Value = marker_name
                Range(MST_MARKER_NUMBER_CELL).Value = mn
                Range(MST_EXPECTED_NUMBER_OF_SCORES).Value = num_criteria * num_assignments
                ' put the focus on the first project to mark
                Range(c2l(MST_FIRST_SCORING_COL) & MST_FIRST_SCORING_ROW).Select
                Range(c2l(MST_FIRST_SCORING_COL) & MST_FIRST_SCORING_ROW).Activate
                ' move the sheet to its own workbook
                With ActiveSheet
                  .Select
                  .Move
                End With
                ' change the name of the sheet to reflect the marker number and name
                ' since names often contain punctuation, keep only character allowed in sheet names, and trim to 30.
                ActiveSheet.Name = GoodTabName(mn & " " & marker_name)
                
                ' fill the scoresheets with made up numbers so they can be loaded into the master sheet
                If simulate_marker_responses Then
                    MakeRandomScores num_assignments
                End If
                
                ' replace the formulas with external references with the resulting text
                ' names of projects
                ConvertCellsDownFromFormula2Text MST_FIRST_SCORING_ROW, MST_NAME_COL
                ConvertRangeToText MST_CRITERIA_NAMES_AND_SCORE_RANGES
                ConvertRangeToText MST_TARGET_SCORING_FRACTION_CELL
                
                ' lock most of the score sheet so the formulas and layout don't get messed up
                If LockMarkerScoresheet = False Then
                    Exit Function
                End If
                ' write out the sheet/workbook and close it, returning focus to the workbook with the macros
                Range(c2l(MST_FIRST_SCORING_COL) & MST_FIRST_SCORING_ROW).Select
                Range(c2l(MST_FIRST_SCORING_COL) & MST_FIRST_SCORING_ROW).Activate
                filestub = DisambiguateFilename(ActiveSheet.Name, mn)
                ActiveWorkbook.SaveAs filename:= _
                    scores_requested_folder & "\" & filestub & scores_ending & ".xlsx", _
                    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
                ActiveWindow.Close
            End If
            num_scoring_files_created = num_scoring_files_created + 1
            marker_workbook_started = False
        End If
        ThisWorkbook.Sheets(MASTER_ASSIGNMENTS_SHEET).Activate
       
    Next mn
        
    'all done. put the focus back where it was before the macro ran
    Sheets(starting_sheet).Activate
    
    If num_markers = num_scoring_files_created Then
        AddMessage "Created " & num_markers & " files for markers to use for scoring."
    Else
        AddMessage num_markers & " possible markers, only found " & num_scoring_files_created & " marking sheets."
    End If
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    CreateAllMarkingSheets = True
End Function
Function ConvertMarkerCommentSheetFormulaToText() As Boolean
    
    ConvertRangeToText SCT_COMPETITION_NAME_CELL
    ConvertRangeToText SCT_MARKER_NAME_CELL
    ConvertRangeToText SCT_PROJECT_NAME_CELL
    ConvertRangeToText SCT_SCORE_CHECK_CELL
    Dim i As Long, name_range As String, min_score_range As String, max_score_range As String
    For i = 1 To max_criteria
        name_range = c2l(Range(SCT_CRITERIA_ONE_NAME_CELL).Column) & _
                    (Range(SCT_CRITERIA_ONE_NAME_CELL).row + SCT_ROWS_PER_CRITERIA * (i - 1))
        ConvertRangeToText name_range
        min_score_range = c2l(Range(SCT_CRITERIA_ONE_MIN_CELL).Column) & _
                    (Range(SCT_CRITERIA_ONE_MIN_CELL).row + SCT_ROWS_PER_CRITERIA * (i - 1))
        ConvertRangeToText min_score_range
        max_score_range = c2l(Range(SCT_CRITERIA_ONE_MAX_CELL).Column) & _
                    (Range(SCT_CRITERIA_ONE_MAX_CELL).row + SCT_ROWS_PER_CRITERIA * (i - 1))
        ConvertRangeToText max_score_range
    Next i
    
    ConvertMarkerCommentSheetFormulaToText = True

End Function

Function SimulateScoresAndComments() As Boolean
    
    ' fill the array with random numbers
    Dim i As Long, first_score_col As Long, first_score_row As Long, score As Double
    Dim lb_cell As String, ub_cell As String, lowerbound As Double, upperbound As Double
    Dim score_range As String, comment_range As String
    first_score_col = Range(SCT_FIRST_CRITERIA_SCORE).Column
    first_score_row = Range(SCT_FIRST_CRITERIA_SCORE).row
    Range(SCT_GENERAL_COMMENT_CELL).Value = "Simply dummy text ... Lorem Ipsum."
    For i = 1 To num_criteria
        lb_cell = c2l(C_FIRST_CRITERIA_MINVALUE_CN) & C_FIRST_CRITERIA_MINVALUE_RN + i - 1
        ub_cell = c2l(C_FIRST_CRITERIA_MINVALUE_CN + 1) & C_FIRST_CRITERIA_MINVALUE_RN + i - 1
        lowerbound = Workbooks(main_workbook).Sheets(CRITERIA_SHEET).Range(lb_cell).Value
        upperbound = Workbooks(main_workbook).Sheets(CRITERIA_SHEET).Range(ub_cell).Value
        score = (upperbound - lowerbound) * Rnd + lowerbound
        score_range = c2l(first_score_col) & (first_score_row + SCT_ROWS_PER_CRITERIA * (i - 1))
        Range(score_range).Value = score
        comment_range = "A" & (first_score_row + 2 + SCT_ROWS_PER_CRITERIA * (i - 1))
        Range(comment_range).Value = "Simply dummy text ... Lorem Ipsum."
    Next i

    SimulateScoresAndComments = True
End Function

Public Function ConvertRangeToText(range_in As String) As Boolean

    Range(range_in).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
    ConvertRangeToText = True

End Function
Public Function LockAndCompressScoresAndCommentsSheet() As Boolean

    'first lock the whole work area
    Range(SCT_FULL_SHEET_RANGE).Select
    Range(FirstCell(SCT_FULL_SHEET_RANGE)).Activate
    Selection.Locked = True
    Selection.FormulaHidden = False
    
    ' now unlock the fields at the top of the sheet
    Range(SCT_COI_RESPONSE_CELL & "," & SCT_CONFIDENCE_LOW_CELL & "," & _
          SCT_CONFIDENCE_MEDIUM_CELL & "," & SCT_CONFIDENCE_HIGH_CELL).Select
    Range(SCT_COI_RESPONSE_CELL).Activate
    Selection.Locked = False
    Selection.FormulaHidden = False
    
    ' unlock the scoring field and the comments field for each criteria
    Dim i As Long, first_score_col As Long, first_score_row As Long
    first_score_col = Range(SCT_FIRST_CRITERIA_SCORE).Column
    first_score_row = Range(SCT_FIRST_CRITERIA_SCORE).row
    Dim unlock_range As String
    unlock_range = SCT_GENERAL_COMMENT_CELL
    For i = 1 To max_criteria
        unlock_range = unlock_range & ","
        unlock_range = unlock_range & _
            c2l(first_score_col) & (first_score_row + SCT_ROWS_PER_CRITERIA * (i - 1))
            unlock_range = unlock_range & ","
        unlock_range = unlock_range & _
                        "A" & ((first_score_row + 2) + SCT_ROWS_PER_CRITERIA * (i - 1)) & _
                        ":K" & ((first_score_row + 2) + SCT_ROWS_PER_CRITERIA * (i - 1))
    Next i
    Range(unlock_range).Select
    Range(FirstCell(unlock_range)).Activate
    Selection.Locked = False
    Selection.FormulaHidden = False
    
    ' hide the rows for the unused criteria
    Dim hide_range As String
    hide_range = (first_score_row - 1 + SCT_ROWS_PER_CRITERIA * num_criteria) & ":" & _
                 (first_score_row - 1 + SCT_ROWS_PER_CRITERIA * max_criteria - 1)
    Range(hide_range).Select
    Selection.EntireRow.Hidden = True
    
    ' apply this access control by locking the sheet
    Range(SCT_COI_RESPONSE_CELL).Select
    Range(SCT_COI_RESPONSE_CELL).Activate
    Dim pwd As String
    pwd = Workbooks(main_workbook).Sheets(SYSTEM_PARAMETERS_SHEET).Range(SP_LOCKED_SHEET_PWD_CELL).Value
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=pwd
    
    LockAndCompressScoresAndCommentsSheet = True
End Function

Public Function GetMarkerName(marker_number As Long) As String
' get the name associated with a marker number, looking it up from the table of marker numbers and names
    
    
    ' Look up the name of a marker from table on the global assignment sheet
    Dim col_num As Long, row_num As Long, marker_row As Long
    
    ' make sure we are on the right sheet, and get the active cell
    Dim current_sheet As String, cell_ref As String
    current_sheet = ActiveSheet.Name
    
    marker_row = FindMarkerRow(marker_number)
    If marker_row <= 0 Then
        GetMarkerName = ""
        Return
    End If
    
    ' make sure we have found the marker number in the table
    col_num = ActiveCell.Column
    row_num = ActiveCell.row
    If (ActiveCell.Value) <> marker_number Then
        MsgBox "[GetMarkerName] Unable to find marker number " & marker_number & " on assignment sheet " & _
                    MASTER_ASSIGNMENTS_SHEET, vbOKOnly
        GetMarkerName = ""
        Return
    End If
    ' move to the right to get the marker's name
    ChangeActiveCell 0, 1
    GetMarkerName = ActiveCell.Value
    
    ' move back to the starting sheet
    Sheets(current_sheet).Activate
    
End Function

Public Function FindMarkerRow(marker_number As Long) As Long 'returns the row of the active marker
    
    Sheets(MARKERS_SHEET).Activate
    ' search in the column of marker numbers for this marker number
    Columns(M_NUMBER_AND_NAME_COLUMNS).Select
    Range("A1").Activate
    If Selection.Find(What:=marker_number, after:=ActiveCell, LookIn:=xlFormulas2, LookAt _
                        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                        False, SearchFormat:=False) Is Nothing Then
        MsgBox "unable to find row for marker " & marker_number, vbCritical
        FindMarkerRow = -1
        Exit Function
    End If
    Selection.Find(What:=marker_number, after:=ActiveCell, LookIn:=xlFormulas2, LookAt _
                        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                        False, SearchFormat:=False).Activate
    FindMarkerRow = ActiveCell.row
    
End Function

Public Function FirstCell(range_string As String) As String
    FirstCell = Left(range_string, InStr(range_string, ":") - 1)
End Function

Public Function LockMarkerScoresheet() As Boolean
    
    ' lock the cells that should not be edited in the scoresheet that the marker will use to score with
    ' these are generally grey-filled
    Const MST_LOCKING_RANGE As String = "A2:C28,D1:O8,I9:O29,D29:H29,G30:I32,J31:M31"
    Range(MST_LOCKING_RANGE).Select
    ' select the first cell named in the range
    Range(FirstCell(MST_LOCKING_RANGE)).Activate
    Selection.FormulaHidden = False
    Selection.Locked = True
    
    ' make sure the scoring entries are not locked.
    Range(c2l(MST_FIRST_SCORING_COL) & MST_FIRST_SCORING_ROW & ":" & _
          c2l(MST_FIRST_SCORING_COL + num_criteria - 1) & _
          (MST_FIRST_SCORING_ROW + max_ass_per_marker - 1)).Select
    Range(c2l(MST_FIRST_SCORING_COL) & MST_FIRST_SCORING_ROW).Activate
    Selection.Locked = False
    
    'now lock the sheet
    Dim pwd As String, this_sheet As String
    this_sheet = ActiveSheet.Name
    pwd = Workbooks(main_workbook).Sheets(SYSTEM_PARAMETERS_SHEET).Range(SP_LOCKED_SHEET_PWD_CELL).Value
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=pwd
    
    LockMarkerScoresheet = True
    
End Function

Public Function old_c2l(col_num As Long) As String
    If Not globals_defined Then
        DefineGlobals
    End If
    If col_num > UBound(col_ltrs) Then
        MsgBox "[c2l]Array col_ltr[] can only handle " & UBound(col_ltrs) & " columns, increase its size", vbOKOnly
    Else
        c2l = col_ltrs(col_num)     'since the array is zero based
    End If
End Function

'Sub test_c2l()
'    DefineGlobals
'
'    Dim num As Long
'    While True
'        num = 10000 * Rnd()
'        MsgBox (c2l(Int(num)))
'    Wend
'End sub

Public Function c2l(col_num As Long) As String

    ' whole different approach - from the stack overflow
    c2l = Split((Columns(col_num).Address(, 0)), ":")(0)
    Exit Function

    
    Dim num_letters As Long, i As Long, num As Long
    Const TEN_TO_MINUS_TEN As Double = 0.0000000001
    
    If (col_num = 0) Then
        MsgBox "[c2l] input must be greater than zero", vbCritical
        c2l = 1
        Exit Function
    End If
    Dim cn As Double, digits(1 To 10) As Double
    cn = col_num
    i = 0
    num_letters = 0
    While cn >= 1
        cn = cn / 26
        num_letters = num_letters + 1
    Wend
    For i = num_letters To 1 Step -1
        num = Int(cn * 26 + TEN_TO_MINUS_TEN)
        If (i > 1) And (Abs(CDbl(num) - cn * 26) < TEN_TO_MINUS_TEN) Then
            i = i - 1
        End If
        c2l = c2l & a2z(num)    ' note a2z goes 0 to 25, and is populated Z, A, B, C, ...
        cn = cn * 26 - num
    Next i
    
End Function

Function DefineGlobals() As Boolean
    ' get some of the parameters, and make sure some global variables are initialized
    
    If (globals_defined) Then
        DefineGlobals = True
        Exit Function
    End If
    
    main_workbook = ThisWorkbook.Name
    
    With ThisWorkbook
        With .Sheets(SYSTEM_PARAMETERS_SHEET)
            max_criteria = .Range(SP_MAX_NUM_OF_CRITERIA_RANGE).Value
            max_markers_per_proj = .Range(SP_MAX_NUMBER_OF_MARKERS_PER_PROJ).Value
            max_ass_per_marker = .Range(SP_MAX_NUMBER_OF_ASSIGNMENTS_PER_MARKER).Value
            max_projects = .Range(SP_MAX_PROJECTS_CELL).Value
'            retain_worksheets = .Range(SP_RETAIN_WORKSHEETS_CELL).Value
            max_keywords = .Range(SP_MAX_KEYWORDS_CELL).Value
            max_markers = .Range(SP_MAX_NUMBER_OF_MARKERS_CELL).Value
            
            'get the strings for selecting appropriate excel files from a folder
            project_expertise_file_pattern = .Range(SP_PROJECT_EXPERTISE_FILE_PATTERN).Value
            keyword_expertise_file_pattern = .Range(SP_KEYWORD_EXPERTISE_FILE_PATTERN).Value
            ss_marks_file_pattern = .Range(SP_MARKER_SCORING_FILE_PATTERN).Value
            ss_marks_comments_file_pattern = .Range(SP_MARKS_WITH_COMMENTS_FILE_PATTERN).Value
            
            ' extract the file ending
            expertise_by_project_ending = Left(project_expertise_file_pattern, _
                                                InStr(project_expertise_file_pattern, ".xlsx") - 1)
            expertise_by_project_ending = Right(expertise_by_project_ending, Len(expertise_by_project_ending) - 1)
            expertise_by_keyword_ending = Left(keyword_expertise_file_pattern, _
                                                InStr(keyword_expertise_file_pattern, ".xlsx") - 1)
            expertise_by_keyword_ending = Right(expertise_by_keyword_ending, Len(expertise_by_keyword_ending) - 1)
            scores_ending = Left(ss_marks_file_pattern, InStr(ss_marks_file_pattern, ".xlsx") - 1)
            scores_ending = Right(scores_ending, Len(scores_ending) - 1)
            scores_with_comments_ending = Left(ss_marks_comments_file_pattern, InStr(ss_marks_comments_file_pattern, ".xlsx") - 1)
            scores_with_comments_ending = Right(scores_with_comments_ending, Len(scores_with_comments_ending) - 1)
            
            same_organization_text = .Range(SP_SAME_ORGANIZATION_TEXT_CELL).Value
            ' get the filename ending that flags a user scoresheet
            simulate_marker_responses = .Range(SP_SIMULATE_MARKER_RESPONSES_CELL).Value
            
        End With
        With .Sheets(COMPETITION_PARAMETERS_SHEET)
            target_ass_per_marker = .Range(CP_TARGET_ASSIGNMENTS_PER_MARKER).Value
            target_markers_per_proj = .Range(CP_TARGET_MARKERS_PER_PROJ).Value
            If target_markers_per_proj > max_markers_per_proj Then
                MsgBox "[DefineGlobals] desired markers per project (" & target_markers_per_proj & _
                        ") exceeds maximum currently possible (" & max_markers_per_proj & ")", vbCritical
                Exit Function
            End If
            root_folder = .Range(CP_COMPETITION_ROOT_FOLDER).Value
            max_first_reader_assignments = .Range(CP_MAX_FIRST_READER_ASSIGNMENTS_CELL).Value
            num_keywords = .Range(CP_NUM_KEYWORDS_CELL).Value
            If (num_keywords > max_keywords) Then
                MsgBox "[DefineGlobals] number of keywords indicated on competition sheet <" & num_keywords & _
                        "> exceeds maximum of keywords supported <" & max_keywords & ">.", vbCritical
                Exit Function
            End If
            normalize_scoring = .Range(CP_USE_NORMALIZED_SCORING_CELL).Value
            
            ' expected folder structure under the root folder
            expertise_by_project_requested_folder = .Range(CP_EXPERTISE_BY_PROJECT_REQUESTED_FOLDER_CELL).Value
            expertise_by_project_received_folder = .Range(CP_EXPERTISE_BY_PROJECT_RECEIVED_FOLDER_CELL).Value
            expertise_by_keyword_requested_folder = .Range(CP_EXPERTISE_BY_KEYWORD_REQUESTED_FOLDER_CELL).Value
            expertise_by_keyword_received_folder = .Range(CP_EXPERTISE_BY_KEYWORD_RECEIVED_FOLDER_CELL).Value
            scores_requested_folder = .Range(CP_SCORES_REQUESTED_FOLDER_CELL).Value
            scores_received_folder = .Range(CP_SCORES_RECEIVED_CELL).Value
            comments_folder = .Range(CP_COMMENTS_FOLDER_CELL).Value
            use_org_disambiguation = .Range(CP_USE_ORG_DISAMBIGUATION_CELL).Value
            use_email_disambiguation = .Range(CP_USE_EMAIL_DISAMBIGUATION_CELL).Value
            gather_comments = .Range(CP_GATHER_COMMENTS_CELL).Value
            output_comments_format = .Range(CP_COMMENT_OUTPUT_FORMAT).Value
        End With
    End With
    
    ' various other variables
    With ThisWorkbook
        num_criteria = .Sheets(CRITERIA_SHEET).Range(C_NUMBER_OF_CRITERIA_CELL).Value
        num_markers = .Sheets(MARKERS_SHEET).Range(M_NUM_MARKERS_CELL).Value     '# of people marking
        num_projects = .Sheets(PROJECTS_SHEET).Range(PS_NUMBER_OF_PROJECTS_CELL).Value '# of projects to mark
        num_keywords = .Sheets(KEYWORDS_SHEET).Range(KW_NUM_KEYWORDS_CELL).Value
        If num_projects > max_projects Then
            MsgBox "This tool set is currently limited to " & max_projects & " projects, " & _
                    num_projects & " in the " & PROJECTS_SHEET & " sheet.", vbCritical
            Exit Function
        End If
        If num_markers > max_markers Then
            MsgBox "This tool set is currently limited to " & max_markers & " markers, " & _
            num_markers & " in the " & MARKERS_SHEET & " sheet.", vbCritical
            Exit Function
        End If
    End With
    
    ' for c2l
    ReDim a2z(0 To 26)
    a2z(0) = "Z"
    Dim i As Long
    For i = 1 To 26
        a2z(i) = Chr(Asc("A") - 1 + i)
    Next i
                    
    Dim letters As Variant
    letters = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", _
                    "T", "U", "V", "W", "X", "Y", "Z", _
                    "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", _
                    "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", _
                    "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", _
                    "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", _
                    "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", _
                    "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", _
                    "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", _
                    "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", _
                    "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", _
                    "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ", _
                    "FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", _
                    "FQ", "FR", "FS", "FT", "FU", "FV", "FW", "FX", "FY", "FZ" _
                    )
    ReDim col_ltrs(1 To UBound(letters))
    For i = 1 To UBound(letters) - 1
        col_ltrs(i) = letters(i - 1)
    Next i
    
    globals_defined = True
    DefineGlobals = True
    
End Function

Public Function ChangeActiveCell(num_row2move As Long, num_col2move As Long) As String 'returns the newly active cell name
    'positive arguments move the active cell to the right and/or down
    Dim curr_row As Long, curr_col As Long
    curr_row = ActiveCell.row
    curr_col = ActiveCell.Column
    Dim curr_cell_name As String
    curr_row = curr_row + num_row2move
    If (curr_row < 1) Or (curr_col + num_col2move < 1) Then
        MsgBox "[ChangeActiveCell] attempt to move before first row or column", vbCritical
        ChangeActiveCell = "A1"
        Range(ChangeActiveCell).Select
    Else
        curr_cell_name = c2l(curr_col + num_col2move) & curr_row
        Range(curr_cell_name).Select
        ChangeActiveCell = curr_cell_name
    End If
End Function

Public Function MakeRandomScores(num_assignments As Long) As Boolean
' fill in the marker scoresheet with random numbers
    
    ' create the random numbers
    Dim upperbound As Long, lowerbound As Long
    Dim i As Long, j As Long, lb_cell As String, ub_cell As String
    Dim scores() As Double
    ReDim scores(1 To num_assignments, 1 To num_criteria)
    
    ' fill the array with random numbers
    For j = 1 To num_criteria
        lb_cell = c2l(C_FIRST_CRITERIA_MINVALUE_CN) & C_FIRST_CRITERIA_MINVALUE_RN + j - 1
        ub_cell = c2l(C_FIRST_CRITERIA_MINVALUE_CN + 1) & C_FIRST_CRITERIA_MINVALUE_RN + j - 1
        lowerbound = Workbooks(main_workbook).Sheets(CRITERIA_SHEET).Range(lb_cell).Value
        upperbound = Workbooks(main_workbook).Sheets(CRITERIA_SHEET).Range(ub_cell).Value
        For i = 1 To num_assignments
            scores(i, j) = (upperbound - lowerbound) * Rnd + lowerbound
        Next i
    Next j
    
    ' now write the array to the expertise column of the expertise sheet
    Dim Destination As Range
    Set Destination = Range(c2l(ActiveCell.Column) & ActiveCell.row)
    Destination.Resize(num_assignments, num_criteria) = scores
'    Destination.Value = Application.Transpose(scores)
        
    Range(c2l(MST_FIRST_SCORING_COL) & MST_FIRST_SCORING_ROW).Select

    MakeRandomScores = False
    
End Function

Public Function CompleteFinalScoresheets() As Boolean 'the is needed to distinguish this sub from the button callback
'   start with a master scoresheet template and populate it from the scoresheets found in a folder

    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If

    Dim starting_sheet As String, starting_workbook As String
    starting_sheet = ActiveSheet.Name
    starting_workbook = ActiveWorkbook.Name
    
    ' load or activate the XLSX workbook containing the master scoresheet
    Dim mss_workbook As String, master_sheet As String, read_range As String
    mss_workbook = ActivateWorkbookBySheetname(MASTER_SCORESHEET)
    If Len(mss_workbook) = 0 Then
        Exit Function
    End If
    
    ' read in the marker and project columns from the shared scoresheet (so we know where to store scores)
    Sheets(SHARED_SCORESHEET).Activate
    read_range = c2l(SS_PROJECT_NUM_COLUMN) & SS_FIRST_DATA_ROW & ":" & _
                 c2l(SS_PROJECT_NUM_COLUMN) & SS_FIRST_DATA_ROW + num_projects * target_markers_per_proj - 1
    ss_project_col = Range(read_range)
    read_range = c2l(SS_MARKER_NUM_COLUMN) & SS_FIRST_DATA_ROW & ":" & _
                 c2l(SS_MARKER_NUM_COLUMN) & SS_FIRST_DATA_ROW + num_projects * target_markers_per_proj - 1
    ss_marker_col = Range(read_range)
        
    ' prepare the master scoresheet for adding scores
    Sheets(MASTER_SCORESHEET).Activate
    master_sheet = ActiveSheet.Name
    ClearMss   'clear the sheet in case it already has scores in it.
    UnsortMasterScoresheet  ' make sure the projects go from 1 to N
            
    ' Ask the user to specify the folder containing the marker score files
    Dim folder_with_scores As String
    ChDir root_folder
    folder_with_scores = SelectFolder("Select folder containing the scoresheets to compile", _
                        root_folder & scores_received_folder)
    If Len(folder_with_scores) = 0 Then
        CompleteFinalScoresheets = False
        Exit Function
    End If

    'load and process the xlsx files that have the right filename pattern to contain scores
    Dim looking As Boolean, marking_sheet As String, marker_num As Long, marker_name As String
    Dim file_name As String, num_ms As Long, num_scores As Long, project_num As Long
    Dim rn As Long, i As Long, j As Long
    Dim num_scores_missing As Long
    Dim file_path As String
    Dim assignment_col As Long, first_col As Long, insert_pos As Long
    Dim file_pattern As String, score_read As String
    ' make space for the comments
    If gather_comments Then
        ReDim comments(1 To num_projects, 1 To num_criteria)
        ReDim general_comments(1 To num_projects)
        ReDim scores(1 To target_ass_per_marker, 1 To num_criteria)
        ReDim project_nums(1 To num_projects, 1 To 1)
        ReDim reader_nums(1 To target_ass_per_marker, 1 To 1)
    End If
    
    ' process the files
    num_ms = 0          ' counter of the number of marking sheets
    looking = True
    While looking
        If num_ms = 0 Then
            If gather_comments Then
                file_name = Dir(folder_with_scores & "\" & ss_marks_comments_file_pattern)
            Else
                file_name = Dir(folder_with_scores & "\" & ss_marks_file_pattern)
            End If
        Else
            file_name = Dir()
        End If
        If Len(file_name) = 0 Then
            looking = False     'no more files to process
        Else
            file_path = folder_with_scores & "\" & file_name
            If gather_comments Then
                marker_num = ReadCommentsAndScores(file_path)
            Else
                marker_num = ReadScoresFromSingleSheet(file_path)
            End If
            
            If marker_num <= 0 Then
                AddMessage "no scores found in " & file_name
            Else
                num_ms = num_ms + 1
                
                'store the scores in the two summary score sheets
                Sheets(master_sheet).Activate
                If AddScoresToMasterSheet(file_path, marker_num) = False Then
                    CompleteFinalScoresheets = False
                    Exit Function
                End If
                Sheets(SHARED_SCORESHEET).Activate
                If AddScoresToSharedSheet(file_path, marker_num) = False Then
                    CompleteFinalScoresheets = False
                    Exit Function
                End If
'                Erase project_nums, reader_nums, scores
            End If
        End If
    Wend
    
    ' all the scoresheets have been read and input, now some final actions
    If SortFinalScoresheets(mss_workbook, master_sheet) = False Then
        Exit Function
    End If
        
    If LabelMasterScoresheetHeaders = False Then
        Exit Function
    End If
    
    ' put the focus on the final scores column and save the workbook
    Sheets(master_sheet).Select
    Sheets(master_sheet).Activate
    Range("A" & MSS_FIRST_PROJECT_ROW & ":" & c2l(MSS_LAST_COL) & MSS_FIRST_PROJECT_ROW).Select
    ActiveWindow.Zoom = True
    Range(c2l(MSS_TOTAL_SCORES_COLUMN) & MSS_FIRST_PROJECT_ROW).Select
    ActiveWorkbook.Save
        
    ' output the evaluators comments if they were loaded
    Dim with_without_comments As String
    If gather_comments Then
        with_without_comments = " with comments"
        If num_ms > 0 Then
            If OutputComments = False Then
                Exit Function
            End If
        End If
    Else
        with_without_comments = "without comments"
    End If
    
    Dim raw_normalized As String
    If normalize_scoring Then
        raw_normalized = "normalized"
    Else
        raw_normalized = "raw"
    End If
    AddMessage "Compiled " & raw_normalized & " scores " & with_without_comments & " from " & num_ms & " markers."
    
    ' put the focus back where it was when the macro started
    Workbooks(starting_workbook).Activate
    Sheets(starting_sheet).Activate
    
    CompleteFinalScoresheets = True
End Function

Function OutputComments() As Boolean
    'save a file for each projects comments
    Dim i As Long, j As Long, rn As Long
    Dim file_name As String
    Dim sheet_name As String
    
    ' copy the template sheet to a new document
    ThisWorkbook.Activate
    sheet_name = DuplicateTemplateSheet(PROJECT_COMMENTS_SHEET)
    With Sheets(sheet_name)
        .Name = SCI_INSTRUCTION_SHEET_NAME
        .Move
    End With
    ' we will repopulate it for each comment sheet
    
    For i = 1 To num_projects
        'populate it with the available comments
        With ActiveSheet
            .Range(PC_PROJECT_NUM_CELL).Value = i
            .Range(PC_PROJECT_NAME_CELL).Value = _
                ThisWorkbook.Sheets(PROJECTS_SHEET).Range(c2l(PS_PROJECT_NAME_COLUMN) & (PS_FIRST_DATA_ROW + i - 1)).Value
            .Range(PC_GENERAL_COMMENTS_CELL).Value = general_comments(i)
            rn = PC_FIRST_CRITERIA_COMMENTS_ROW
            For j = 1 To num_criteria
                .Range("A" & rn).Value = comments(i, j)
                rn = rn + PC_NUM_ROWS_PER_COMMENT
            Next j
            
            ' save it to a format accessible by a word-processor
            Dim filestub As String, ending As String, saveas_type As Long
            
            Select Case output_comments_format
            Case "PRN"
                saveas_type = xlTextPrinter
                ending = ".prn"
            Case "HTML"
                saveas_type = xlHtml
                ending = ".htm"
            Case "XLSX"
                saveas_type = xlOpenXMLWorkbook
                ending = ".xlsx"
            Case "TEXT"
                saveas_type = xlTextWindows
                ending = ".txt"
            Case Else
                MsgBox "unknown file output format " & output_comments_format, vbCritical
                Exit Function
            End Select
            filestub = GoodTabName(i & " " & ActiveSheet.Range(PC_PROJECT_NAME_CELL).Value)
            file_name = root_folder & comments_folder & filestub & ending
            ActiveWorkbook.SaveAs filename:=file_name _
                , FileFormat:=saveas_type, ReadOnlyRecommended:=False, CreateBackup:=False
        End With
    Next i
    
    ' delete the template workbook
    ActiveWorkbook.Close
    
    OutputComments = True
End Function
    

Function SortFinalScoresheets(mss_workbook As String, master_sheet As String) As Boolean
    'sort both versions of the final score sheets by decreasing scores
    
    Workbooks(mss_workbook).Activate
    Sheets(SHARED_SCORESHEET).Activate
    Dim sort_range As String, sort_key As String
    sort_key = c2l(SS_FINAL_TOTAL_SCORES_COLUMN) & SS_FIRST_DATA_ROW & ":" & _
               c2l(SS_FINAL_TOTAL_SCORES_COLUMN) & (SS_FIRST_DATA_ROW + num_projects - 1)
    sort_range = c2l(SS_FINAL_PROJ_COLUMN) & SS_FIRST_DATA_ROW & ":" & _
               c2l(SS_FINAL_TOTAL_SCORES_COLUMN) & (SS_FIRST_DATA_ROW + num_projects - 1)
    Range(sort_range).Select
    Range(FirstCell(sort_key)).Activate
    ActiveWorkbook.Worksheets(SHARED_SCORESHEET).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(SHARED_SCORESHEET).Sort.SortFields.Add2 Key:= _
        Range(sort_key), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(SHARED_SCORESHEET).Sort
        .SetRange Range(sort_range)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Sheets(master_sheet).Activate
    sort_key = c2l(MSS_TOTAL_SCORES_COLUMN) & MSS_FIRST_PROJECT_ROW & ":" & _
               c2l(MSS_TOTAL_SCORES_COLUMN) & (MSS_FIRST_PROJECT_ROW + num_projects - 1)
    sort_range = "A" & MSS_FIRST_PROJECT_ROW & ":" & _
               c2l(MSS_LAST_COL + 2) & (MSS_FIRST_PROJECT_ROW + num_projects - 1)
    Range(sort_range).Select
    Range(FirstCell(sort_key)).Activate
    ActiveWorkbook.Worksheets(master_sheet).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(master_sheet).Sort.SortFields.Add2 Key:= _
        Range(sort_key), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(master_sheet).Sort
        .SetRange Range(sort_range)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
    SortFinalScoresheets = True
End Function
    
Function LabelMasterScoresheetHeaders() As Boolean
    ' put in the correct header (normalized or not)
    Dim stub As String, i As Long
    Dim title_range As String
    If normalize_scoring Then
        stub = "normalized"
    Else
        stub = "raw"
    End If
    stub = stub & ") criteria scores"
    For i = 1 To max_markers_per_proj
        title_range = c2l(MSS_FIRST_SCORE_COL + 1 + (i - 1) * (max_criteria + 2)) & MSS_MARKER_NUMBER_ROW
        Range(title_range).Value = "Marker #" & i & " (" & stub
    Next i
    LabelMasterScoresheetHeaders = True
End Function

Function AddScoresToSharedSheet(file_path As String, marker_num As Long) As Boolean
    'this function assumes the shared scoresheet is active
    Dim i As Long, pn As Long, j As Long, k As Long
    Dim insert_start As String, row() As Variant
    ReDim row(1 To num_criteria)
    Dim Destination As Range
    For i = 1 To UBound(ss_marker_col, 1)
        If ss_marker_col(i, 1) = marker_num Then
            'we've found the first row for this marker
            j = i
            pn = 1
            ' look through the remaining project rows for one that match the marker's project assignments
            While pn <= UBound(scores, 1)
                If (ss_marker_col(i, 1) = marker_num) And (ss_project_col(j, 1) = project_nums(pn, 1)) Then
                    ' we have a project row that was assigned to this marker
                    For k = 1 To num_criteria   ' gather the row of scores on this project
                        row(k) = scores(pn, k)
                    Next k
                    ' copy the scores into the table
                    insert_start = c2l(SS_FIRST_RAW_COLUMN) & (SS_FIRST_DATA_ROW + j - 1)
                    Set Destination = Range(insert_start)
                    Set Destination = Destination.Resize(1, UBound(row))
                    Destination.Value = row
                    pn = pn + 1 'look for a match to the next project assigned to the marker
                Else
                    If j > UBound(ss_marker_col, 1) Then
                        MsgBox "AddScoresToSharedSheet error - did not find project " & project_nums(pn, 1) & _
                                " for marker " & marker_num, vbCritical
                        Exit Function
                    Else
                        j = j + 1   ' see if the next row in the table is for this project and marker
                    End If
                End If
            Wend 'pn
            ' we've stored the scores from this marker, exit
            Exit For
        End If
    Next i
    
    AddScoresToSharedSheet = True
    
End Function

Private Function ReadCommentsAndScores(file_path As String) As Long
    'open the file with a scoresheet
    Dim marker_num As Long, num_scores As Long, i As Long, j As Long, num_assigned_as_long
    Dim num_assigned As Long, row_num As Long
    Dim tab_names() As String, comment As Variant
    Workbooks.Open filename:=file_path
    
    ' load in the structure information for the book from the table on the instructions sheet
    num_assigned = Sheets(SCI_INSTRUCTION_SHEET_NAME).Range("B" & SCI_PROJECT_COUNT_ROW).Value
'    ReDim proj_nums(1 To num_assigned)
    ReDim tab_names(1 To num_assigned)
    For i = 1 To num_assigned
        project_nums(i, 1) = Sheets(SCI_INSTRUCTION_SHEET_NAME).Range("A" & (SCI_PROJECT_COUNT_ROW + 1 + i)).Value
        tab_names(i) = Sheets(SCI_INSTRUCTION_SHEET_NAME).Range("C" & (SCI_PROJECT_COUNT_ROW + 1 + i)).Value
    Next i

    ' go through each of the marks/comments sheets and extract the scores and comments info
    For i = 1 To num_assigned
        With Sheets(tab_names(i))
            row_num = 2
            reader_nums(i, 1) = .Range(SCT_READER_NUM_CELL).Value
            AppendComment CLng(reader_nums(i, 1)), .Range(SCT_GENERAL_COMMENT_CELL).Value, general_comments(project_nums(i, 1))
            row_num = Range(SCT_FIRST_CRITERIA_SCORE).row
            For j = 1 To num_criteria
                scores(i, j) = .Range(c2l(SCT_SCORE_COLUMN) & row_num).Value
                ' get the comment (if provided and append it to the comments for that criteria
                AppendComment CLng(reader_nums(i, 1)), .Range("A" & (row_num + 2)).Value, comments(project_nums(i, 1), j)
                row_num = row_num + SCT_ROWS_PER_CRITERIA
            Next j
        End With
    Next i
    
    ReadCommentsAndScores = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, " ") - 1)
    
    ActiveWorkbook.Close
    
 End Function

Private Function AppendComment(reader_num As Long, comment_read As String, comments As Variant) As Boolean
    If Len(CStr(comments)) > 1 Then
        comments = comments & vbCrLf
    End If
    comments = comments & "Comment from reader #" & reader_num & ":" & vbCrLf & comment_read
    AppendComment = True
End Function

Private Function ReadScoresFromSingleSheet(file_path As String) As Long
    ' returns marker number who did the scoring
    
    'open the file with a scoresheet
    Dim marker_num As Long, num_scores As Long
    Workbooks.Open filename:=file_path
    ' get name of scoring sheet (should be the only sheet)
    If Sheets.Count > 1 Then
        MsgBox "[ReadScoresFromSingleSheet] Expected only one sheet in book, found " & Sheets.Count, vbCritical
        Exit Function
    End If
    marker_num = ActiveSheet.Range(MST_MARKER_NUMBER_CELL).Value
    
    ' Extract the scores and associated data from the sheet
    num_scores = Range(MST_PROJECT_COUNT_CELL).Value
    Dim first_score_column As Long
    Dim projects_range As String, readers_range As String, scores_range As String
    'load the project numbers assigned to this marker
    projects_range = MST_PROJECT_NUM_COL & MST_FIRST_SCORING_ROW & ":" & _
                     MST_PROJECT_NUM_COL & MST_FIRST_SCORING_ROW + (num_scores - 1)
    project_nums = Range(projects_range)
    'load the reader numbers for each assignment
    readers_range = MST_READER_NUM_COL & MST_FIRST_SCORING_ROW & ":" & _
                    MST_READER_NUM_COL & MST_FIRST_SCORING_ROW + (num_scores - 1)
    reader_nums = Range(readers_range)
    If normalize_scoring Then
        first_score_column = MST_FIRST_NORMALIZED_SCORE_COLUMN
    Else
        first_score_column = MST_FIRST_SCORING_COL
    End If
    scores_range = c2l(first_score_column) & MST_FIRST_SCORING_ROW & ":" & _
                    c2l(first_score_column + num_criteria - 1) & (MST_FIRST_SCORING_ROW + num_scores - 1)
    scores = Range(scores_range)
    ActiveWorkbook.Close
    
    ReadScoresFromSingleSheet = marker_num
    
End Function

Function AddScoresToMasterSheet(file_path As String, marker_num As Long) As Boolean
    ' this function assumes the master scoresheet is selected
    
    ' for loading the scores into the master sheet
    Dim num_scores As Long, num_scores_missing, i As Long, j As Long
    num_scores = UBound(scores, 1)
    Dim scores_row() As Double
    ReDim scores_row(1 To num_criteria)
    For i = 1 To num_scores
        num_scores_missing = 0
        For j = 1 To num_criteria
            If Len(scores(i, j)) > 0 Then
                scores_row(j) = CDbl(scores(i, j))
            Else
                num_scores_missing = num_scores_missing + 1
            End If
        Next j
        If num_scores_missing > 0 Then
            AddMessage "Project " & project_nums(i, 1) & " in file <" & file_path & "> is missing " & _
                        num_scores_missing & " score(s)."
        Else
            'find the project that corresponds to this row of scores
            Range(c2l(MSS_PROJECT_COLUMN) & MSS_FIRST_PROJECT_ROW & ":" & _
                    c2l(MSS_PROJECT_COLUMN) & MSS_FIRST_PROJECT_ROW + num_projects - 1).Select
            ' was xlformulas2
            Selection.Find(What:=project_nums(i, 1), after:=ActiveCell, LookIn:=xlValues, _
                    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False).Activate
            'insert the scores
            If (InsertScores(CLng(reader_nums(i, 1)), marker_num, scores_row) = False) Then
                Exit Function
            End If
        End If
    Next i
    
    AddScoresToMasterSheet = True
    
End Function
Function ActivateWorkbookBySheetname(sheet_name As String) As String ' returns the name of the workbook
    ' activate or open the workbook containing the master scoresheet

    'loop through the open workbooks looking for an XLSX file containing a sheet
    ' whose name is given by the constant MASTER_SCORESHEET
    Dim i As Long, j As Long
    Dim found As Boolean
    For i = 1 To Workbooks.Count
        With Workbooks(i)
            If Right(.FullName, 4) = "xlsx" Then
                For j = 1 To .Sheets.Count
                    If .Sheets(j).Name = sheet_name Then
                        found = True
                        .Activate
                        .Sheets(j).Activate
                        ActivateWorkbookBySheetname = ActiveWorkbook.Name
                        Exit Function
                    End If
                Next j
            End If
        End With
    Next i
    
    ' the one we need is not among the active workbooks, ask the user to select a file
    Dim looking As Boolean
    looking = True
    While looking
        Dim file_name As String
        file_name = Application.GetOpenFilename(FileFilter:="Excel files (*.xlsx*), *.xlsx", _
                                title:="Choose the Excel file to update its master scoresheet", MultiSelect:=False)
        If (Len(file_name) = 0) Or (file_name = "False") Then
            Exit Function
        End If
        
        ' open the selected file
        Workbooks.Open filename:=file_name
        
        ' check that it contains the desired sheet
        With ActiveWorkbook
            For j = 1 To .Sheets.Count
                If .Sheets(j).Name = sheet_name Then
                    looking = False
                    ActivateWorkbookBySheetname = ActiveWorkbook.Name
                    Exit Function
                End If
            Next j
        End With
        ' sheet not found, let the user know
        Dim msg As String
        msg = " workbook " & file_name & " does not contain a sheet named " & sheet_name & _
                ". Please select a file that does"
        If MsgBox(msg, vbOKCancel) = vbCancel Then
            ActivateWorkbookBySheetname = ""
            Exit Function
        End If
    Wend
    
End Function

Function InsertScores(assignment_col As Long, marker_num As Long, scores_row() As Double) As Boolean
    ' insert the three (normalized) scores in the master sheet, and the marker who made them,
    ' assuming we are on the correct row.
    If marker_num < 1 Or marker_num > num_markers Then
        MsgBox "InsertScores] unexpected marker_num: " & marker_num, vbCritical
        InsertScores = False
        Exit Function
    End If
    Dim start_col As Long, i As Long
    start_col = MSS_FIRST_SCORE_COL + (max_criteria + 2) * (assignment_col - 1)
    Range(c2l(start_col - 1) & ActiveCell.row).Value = marker_num
    For i = 1 To UBound(scores_row)
        Range(c2l(start_col + i - 1) & ActiveCell.row).Value = scores_row(i)
    Next i
    InsertScores = True
    
End Function


Public Function MSSCOnvertFormulasToText()
'
'   replace the vlookups for the marker # columns, and the marker name lookups
'   with their text equivalents, so that the MSS does not have links to other workbooks
    ConvertCellsDownFromFormula2Text MSS_FIRST_PROJECT_ROW, "B"
    ConvertCellsDownFromFormula2Text MSS_FIRST_PROJECT_ROW, "C"
    ConvertCellsDownFromFormula2Text MSS_FIRST_PROJECT_ROW, "J"
    ConvertCellsDownFromFormula2Text MSS_FIRST_PROJECT_ROW, "Q"
    ConvertCellsDownFromFormula2Text MSS_FIRST_PROJECT_ROW, "X"
    ConvertCellsDownFromFormula2Text MSS_FIRST_PROJECT_ROW, "AE"
    
    ' convert the vlookup for the criteria and scoring ranges to text
    ConvertRangeToText "D4:H4"

End Function

Public Function ConvertCellsDownFromFormula2Text(start_row As Long, column_name As String)
    Dim last_row As Long
    last_row = start_row + max_projects - 1

    Range(column_name & start_row & ":" & column_name & last_row).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

End Function
    
Public Function SelectFolder(title As String, start_folder As String) As String
    Dim diaFolder As FileDialog
    Dim selected As Boolean, entry_dir As String

'    entry_dir = CurDir
'    ChDir (start_folder)
    ' Open the file dialog
    
    
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    
    diaFolder.AllowMultiSelect = False
    diaFolder.title = title
    Dim path As String
    path = Dir(start_folder, vbDirectory)   ' check if the folder exists (return will be non-null)
    If Len(path) = 0 Then
        MsgBox "WARNING: File path in [SelectFolder] <" & start_folder & "> does not exist", vbOK
    Else
        If Len(start_folder) > 0 Then
            diaFolder.InitialFileName = start_folder
        End If
    End If
    
    selected = diaFolder.Show

    If selected Then
       SelectFolder = diaFolder.SelectedItems(1)
    End If
'    ChDir (entry_dir)

    Set diaFolder = Nothing
End Function

Public Function getMarkingSheetName(file_name As String, name_suffix As String, _
                                    marker_num As Long, marker_name As String) _
                                    As Long
' file_name = "18 Sigurer.xlsx"
' sheet_name = "18 Sigurer"
' marker_num = 18
' Marker name = "Sigurer"  (not the exact maker's name as accents and punctuation were removed)

    Dim i As Long, sheet_name As String
    Dim number_from_sheet_name As String, number_from_file_name As String
    number_from_file_name = Left(file_name, InStr(file_name, " ") - 1)      ' file_name leads with a number
    For i = 1 To ActiveWorkbook.Sheets.Count        ' make sure the book name starts with a number
        sheet_name = Sheets(i).Name
        If InStr(Sheets(i).Name, " ") > 0 Then      ' and has to have a space in the sheet name
            number_from_sheet_name = Left(Sheets(i).Name, InStr(Sheets(i).Name, " ") - 1)
            If number_from_file_name = number_from_sheet_name Then
                marker_num = Val(number_from_file_name)
                sheet_name = Left(file_name, Len(file_name) - 5) ' strip off the extent
                marker_name = Right(file_name, Len(file_name) - InStr(file_name, " "))   ' get string without the number
                marker_name = Left(marker_name, InStr(marker_name, name_suffix) - 1)    ' remove the suffix to get the name
                getMarkingSheetName = i
                Exit Function
            End If
        End If
    Next i
    MsgBox "ERROR getting marker number and name from file " & file_name, vbCritical
    
End Function

Public Function GetFileSaveasName(dialog_title As String, initial_name As String) As String
    Dim varResult As Variant
    'displays the save file dialog
    varResult = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", _
       title:=dialog_title, InitialFileName:=SAVE_FOLDER & initial_name)
    'checks to make sure the user hasn't canceled the dialog
    If varResult <> False Then
        Exit Function
    Else
        GetFileSaveasName = varResult
    End If
End Function

Public Function LoadMarkerProjectExpertiseIntoPXM() As Boolean
    ' load the data from all the project expertise sheets in a folder into the PXM table
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    ' initialize the PXM sheet
    Dim PXM_workbook As String
    PXM_workbook = ActiveWorkbook.Name
    Sheets(PROJECT_X_MARKER_SHEET).Select
    
    ClearPXMSheet
    
    ' load the data from the Project Keyword, Marker Expertise, Calculated PXM
    
    ' get the folder name with the expertise files to load
    Dim folder_with_expertises As String
'    ChDir root_folder
    folder_with_expertises = SelectFolder("Select folder containing the expertise about projects of potential markers", _
                root_folder & expertise_by_project_received_folder)
    If Len(folder_with_expertises) = 0 Then
        Exit Function
    End If

    'load and process the xlsx files that have the right filename pattern.
    Dim looking As Boolean, expertise_sheet As String, marker_num As Long, marker_name As String
    Dim file_name As String, num_ms As Long, num_expertises As Long, project_num As Long
    Dim rn As Long, i As Long, j As Long, num_es As Long, insert_pos As Long
    Dim first_col As Long
    num_es = 0                      ' counter of the number of expertise sheets
    Dim expertise() As Variant      ' arrays read from the expertise sheet
    Dim COIs() As Variant           ' the contents will be strings, but use variants to get from the S/S
    Dim finding As Boolean
    Dim PXM_range As String
    Dim expertise_out() As Variant
    ReDim expertise_out(1 To num_projects)
    Dim expertise_workbook As String
    
    finding = True
    While finding
        If num_es = 0 Then
            file_name = Dir(folder_with_expertises & "\" & project_expertise_file_pattern)
        Else
            file_name = Dir()
        End If
        If Len(file_name) = 0 Then
            finding = False     'no more files to process
        Else        'open the file with a expertise sheet
            ' get name of expertise sheet from filename
            Workbooks.Open filename:=folder_with_expertises & "\" & file_name
            expertise_workbook = ActiveWorkbook.Name
            Dim sht_num As Long
            sht_num = getMarkingSheetName(file_name, expertise_by_project_ending, marker_num, marker_name)
            expertise_sheet = Sheets(sht_num).Name
  'PROJECT VERSION
            ' load the keyword expertise column from that sheet
            COIs = Range(c2l(MPET_COI_COLUMN) & MPET_FIRST_DATA_ROW & ":" & _
                         c2l(MPET_COI_COLUMN) & (MPET_FIRST_DATA_ROW + num_projects - 1))
            
            ' load the project COI column from the COI sheet
            expertise = Range(c2l(MPET_EXPERTISE_COLUMN) & MPET_FIRST_DATA_ROW & ":" & _
                                c2l(MPET_EXPERTISE_COLUMN) & (MPET_FIRST_DATA_ROW + num_projects - 1))
  
            ' close the expertise workbook
            Workbooks(expertise_workbook).Close
            
            'update the expertise to exclude rows with COI signalled
            For i = 1 To num_projects
                Select Case COIs(i, 1)
                Case "X", "Y", "Mentor", "MENTOR", same_organization_text
                    expertise_out(i) = "X"
                Case "N", ""
                    expertise_out(i) = LMH2Percent(CStr(expertise(i, 1)))
                Case Else
                    MsgBox "[LoadMarkerProjectExpertiseIntoPXM] unexpected COI(" & i & ") value: " & COIs(i), vbCritical
                    Exit Function
                End Select
            Next i
            'find the column for this marker
            Sheets(PROJECT_X_MARKER_SHEET).Select
            PXM_range = c2l(PXM_FIRST_PXM_COL) & PXM_MARKER_NUM_ROW & ":" & _
                        c2l(PXM_FIRST_PXM_COL + num_markers - 1) & PXM_MARKER_NUM_ROW
            Range(PXM_range).Select 'select the row containing the marker #'s
            If Selection.Find(What:=marker_num, LookIn:=xlFormulas2, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False) Is Nothing Then
                'nothing found, nothing for this marker so exit the loop
                MsgBox _
                  "[LoadMarkerProjectExpertiseIntoPXM] error finding the marker numbers - check it is the right sheet", _
                  vbCritical
                finding = False
                Exit Function
            End If
            Selection.Find(What:=marker_num, after:=ActiveCell, LookIn:=xlFormulas2, _
            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
            
            ' insert the column of data about the marker's expertise and availability into the crosswalk sheet
            Range(c2l(ActiveCell.Column) & PXM_FIRST_DATA_ROW).Select
            Dim Destination As Range
            Set Destination = ActiveCell
            Set Destination = Destination.Resize(UBound(expertise_out), 1)
            Destination.Value = Application.Transpose(expertise_out)
            num_es = num_es + 1
            If retain_worksheets Then
                MsgBox "need to add code to retain markers' worksheets", vbCritical
            End If
        End If
    Wend
    
    ' save the workbook containing the PXM sheet
    Range(c2l(PXM_FIRST_PXM_COL) & (PXM_FIRST_DATA_ROW)).Select
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Dim msg As String
    msg = "Loaded " & num_es & " expertise and conflict profiles into PXM table. "
    AddMessage msg
    
    LoadMarkerProjectExpertiseIntoPXM = True
    
End Function

Public Function LoadMarkerKeywordExpertiseIntoPXM() As Boolean
    ' load the data from all the project expertise sheets in a folder into the PXM table
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    ' initialize the PXM sheet
    Dim PXM_workbook As String
    PXM_workbook = ActiveWorkbook.Name
    Sheets(PROJECT_X_MARKER_SHEET).Select
    
    ClearPXMSheet
    
    ' get the folder name with the expertise files to load
    Dim folder_with_expertises As String
'    ChDir root_folder
    folder_with_expertises = SelectFolder("Select folder containing the confidence of markers by Keywords", _
                root_folder & expertise_by_keyword_received_folder)
    If Len(folder_with_expertises) = 0 Then
        Exit Function
    End If

    'load and process the xlsx files that have the right filename pattern.
    Dim looking As Boolean, expertise_sheet As String, marker_num As Long, marker_name As String
    Dim file_name As String, num_ms As Long, num_expertises As Long, project_num As Long
    Dim rn As Long, i As Long, j As Long, num_es As Long, insert_pos As Long
    Dim first_col As Long
    num_es = 0                      ' counter of the number of expertise sheets
    Dim expertise() As Variant      ' arrays read from the expertise sheet
    Dim COIs() As Variant           ' the contents will be strings, but use variants to get from the S/S
    Dim finding As Boolean
    Dim keyword_row_range As String
    Dim expertise_out() As Variant
    ReDim expertise_out(1 To num_projects)
    Dim expertise_workbook As String
    ReDim competition_COIs(1 To num_projects, 1 To num_markers)
    
    finding = True
    While finding
        If num_es = 0 Then
            file_name = Dir(folder_with_expertises & "\" & keyword_expertise_file_pattern)
        Else
            file_name = Dir()
        End If
        If Len(file_name) = 0 Then
            finding = False     'no more files to process
        Else        'open the file with a expertise sheet
            ' get name of expertise sheet from filename
            Workbooks.Open filename:=folder_with_expertises & "\" & file_name
            expertise_workbook = ActiveWorkbook.Name
            Dim sht_num As Long
            sht_num = getMarkingSheetName(file_name, expertise_by_keyword_ending, marker_num, marker_name)
            expertise_sheet = Sheets(sht_num).Name

'KEYWORD version
            ' load the keyword expertise column from that sheet
            Sheets(expertise_sheet).Select
            expertise = Range(MKET_EXPERTISE_COLUMN & MKET_FIRST_DATA_ROW & ":" & _
                              MKET_EXPERTISE_COLUMN & (MKET_FIRST_DATA_ROW + num_keywords - 1))
            
            ' load the project COI column from the COI sheet
            Sheets(MKET_COI_SHEET_NAME).Select
            COIs = Range(c2l(MPET_COI_COLUMN) & MPET_FIRST_DATA_ROW & ":" & _
                         c2l(MPET_COI_COLUMN) & (MPET_FIRST_DATA_ROW + num_projects - 1))
            'load this marker's set of project COIs into the project x marker array
            For i = 1 To num_projects
                If Len(COIs(i, 1)) > 0 Then
                    competition_COIs(i, marker_num) = "X"
                Else
                    ' empty cells mean no COI
                End If
            Next i
            ' close the expertise workbook
            Workbooks(expertise_workbook).Close
            
            ' select the destination sheet
            Sheets(MARKER_EXPERTISE_SHEET).Select
            'find the column for this marker
            keyword_row_range = c2l(ME_FIRST_MARKER_DATA_COL) & (ME_FIRST_MARKER_DATA_ROW + marker_num - 1) & ":" & _
                                c2l(ME_FIRST_MARKER_DATA_COL + num_keywords - 1) & (ME_FIRST_MARKER_DATA_ROW + marker_num - 1)
            ' insert the data about the marker's expertise as a row in the table
            Dim Destination As Range
            Set Destination = Range(keyword_row_range)
            Destination.Resize(UBound(expertise, 2), UBound(expertise, 1)).Value = Application.Transpose(expertise)
            num_es = num_es + 1
            If retain_worksheets Then
                MsgBox "need to add code to retain markers' worksheets", vbCritical
            End If
        End If
    Wend
    
    ' write the array of competition COIs to the PXM sheet
    Sheets(PROJECT_X_MARKER_SHEET).Select
    Dim cf_range As String
    cf_range = c2l(PXM_FIRST_PXM_COL) & PXM_FIRST_DATA_ROW & ":" & _
                c2l(PXM_FIRST_PXM_COL + num_markers - 1) & (PXM_FIRST_DATA_ROW + num_projects - 1)
    Set Destination = Range(cf_range)
    Destination.Resize(UBound(competition_COIs, 1), UBound(competition_COIs, 2)) = competition_COIs
    
    ' save the workbook containing the PXM sheet
    Range(c2l(PXM_FIRST_PXM_COL) & (PXM_FIRST_DATA_ROW)).Select
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Dim msg As String
    msg = "Loaded " & num_es & " expertise and conflict profiles into PXM table. "
    AddMessage msg
    
    LoadMarkerKeywordExpertiseIntoPXM = True
    
End Function

Public Function LoadMarkerExpertiseIntoCrosswalk() As Boolean
    'make a duplicate of the template sheet
    'load the list of projects to mark
    'for each marker load the expertises they signaled, paying attention to any COIs signaled
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    ' initialize the crosswalk sheet
    Dim crosswalk_sheet As String
    ' this approach adds the data to the sheet in the current workbook
    crosswalk_sheet = EXPERTISE_CROSSWALK_SHEET
    
    clear_crosswalk_sheet (crosswalk_sheet)
    
    Dim crosswalk_workbook As String
    crosswalk_workbook = ActiveWorkbook.Name
    
    ' get the folder name with the expertise files to load
    Dim folder_with_expertises As String
'    ChDir root_folder
    folder_with_expertises = SelectFolder("Select folder containing the expertise specified by potential markers", _
                root_folder & expertise_by_project_received_folder)
    If Len(folder_with_expertises) = 0 Then
        Exit Function
    End If

    'load and process the xlsx files that have the right filename pattern.
    Dim looking As Boolean, expertise_sheet As String, marker_num As Long, marker_name As String
    Dim file_name As String, num_ms As Long, num_expertises As Long, project_num As Long
    Dim rn As Long, i As Long, j As Long, num_es As Long, insert_pos As Long
    Dim first_col As Long
    num_es = 0                      ' counter of the number of expertise sheets
    Dim expertise() As Variant      ' arrays read from the expertise sheet
    Dim COIs() As Variant           ' the contents will be strings, but use variants to get from the S/S
    Dim finding As Boolean
    Dim marker_number_row As String
    Dim expertise_out() As Variant
    ReDim expertise_out(1 To num_projects)
    Dim expertise_workbook As String
    
    finding = True
    While finding
        If num_es = 0 Then
            file_name = Dir(folder_with_expertises & "\" & expertise_file_pattern)
        Else
            file_name = Dir()
        End If
        If Len(file_name) = 0 Then
            finding = False     'no more files to process
        Else        'open the file with a expertise sheet
            Workbooks.Open filename:=folder_with_expertises & "\" & file_name
            expertise_workbook = ActiveWorkbook.Name
            
            ' get name of expertise sheet from filename
            Dim sht_num As Long
            sht_num = getMarkingSheetName(file_name, EXPERTISE_ENDING, marker_num, marker_name)
            expertise_sheet = Sheets(sht_num).Name
            
            'move it into the workbook
            Sheets(expertise_sheet).Select
            num_es = num_es + 1
            insert_pos = Workbooks(crosswalk_workbook).Sheets.Count
            Sheets(expertise_sheet).Move after:=Workbooks(crosswalk_workbook).Sheets(insert_pos)
            Sheets(expertise_sheet).Activate
            Workbooks(expertise_workbook).Close
            
            'copy the column of expertise and the column of COI information from the expertise sheet
            expertise = Range(c2l(MPET_EXPERTISE_COLUMN) & "2:" & c2l(MPET_EXPERTISE_COLUMN) & (2 + num_projects - 1))
            COIs = Range(c2l(MPET_COI_COLUMN) & MPET_FIRST_DATA_ROW & ":" & _
                         c2l(MPET_COI_COLUMN) & (MPET_FIRST_DATA_ROW + num_projects - 1))
            'update the expertise to exclude rows with COI signalled
            For i = 1 To num_projects
                Select Case COIs(i, 1)
                Case "X", "Y", "Mentor", "MENTOR", same_organization_text
                    expertise_out(i) = "X"
                Case "N", ""
                    expertise_out(i) = LMH2Percent(CStr(expertise(i, 1)))
                Case Else
                    MsgBox "[load_marker_expertise_in_crosswalk] unexpected COI[" & i & "] value: " & COIs(i), vbCritical
                    Exit Function
                End Select
            Next i
            'find the column for this marker
            Sheets(crosswalk_sheet).Select
            marker_number_row = c2l(EC_DATA_FIRST_MARKER_COLUMN) & _
                                    (EC_DATA_FIRST_MARKER_ROW - 1) & ":" & _
                                c2l(EC_DATA_FIRST_MARKER_COLUMN - 1 + num_markers) & _
                                    (EC_DATA_FIRST_MARKER_ROW - 1)
            Range(marker_number_row).Select 'select the row containing the marker #'s
            If Selection.Find(What:=marker_num, LookIn:=xlFormulas2, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False) Is Nothing Then
                'nothing found, nothing for this marker so exit the loop
                MsgBox _
                  "[LoadMarkerExpertiseIntoCrosswalk] error finding the marker numbers - check it is the right sheet", _
                  vbCritical
                finding = False
                Exit Function
            End If
            Selection.Find(What:=marker_num, after:=ActiveCell, LookIn:=xlFormulas2, _
            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
            
            ' insert the column of data about the marker's expertise and availability to the crosswalk sheet
            ChangeActiveCell 1, 0  ' move down to the first row of data
            Dim Destination As Range
            Set Destination = ActiveCell
            Set Destination = Destination.Resize(UBound(expertise_out), 1)
            Destination.Value = Application.Transpose(expertise_out)
            Sheets(expertise_sheet).Delete
        End If
    Wend
    
    ' save the workbook containing the expertise crosswalk sheet and the loaded expertise inputs
    Range(c2l(MPET_COI_COLUMN) & MPET_FIRST_DATA_ROW).Select
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Dim msg As String
    msg = "Loaded " & num_es & " expertise and conflict profiles into crosswalk table. "
    AddMessage msg
    
    LoadMarkerExpertiseIntoCrosswalk = True
    
End Function

'Public Function LMH2PercentSS(letter As String) As Double
'
'    Select Case letter
'    Case "H"
'        LMH2PercentSS = 1
'    Case "M"
'        LMH2PercentSS = TWO_THIRDS
'    Case "L"
'        LMH2PercentSS = ONE_THIRD
'    Case ""
'        LMH2PercentSS = 0
'    Case Else
'        MsgBox "Unexpected letter to LMH2Percent <" & letter & ">", vbCritical
'    End Select
'End Function

    

Public Function LMH2Percent(letter As String) As Double
    Select Case letter
    Case "H"
        LMH2Percent = 1
    Case "M"
        LMH2Percent = TWO_THIRDS
    Case "L", ""    ' no expertise provided = Low
        LMH2Percent = ONE_THIRD
    Case "X"
        ' ignore conflicts
    Case Else
        MsgBox "Unexpected letter to LMH2Percent <" & letter & ">", vbCritical
    End Select
    
End Function

Public Function AssignMarkersBasedOnConfidence() As Boolean
    
    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    Dim crosswalk_book As String, starting_sheet As String
    crosswalk_book = ActiveWorkbook.Name
    starting_sheet = ActiveSheet.Name

    ' read the PXM data from the PXM sheet into the marker confidence array
    Sheets(PROJECT_X_MARKER_SHEET).Select
    Dim PXM_range As String, mc_range As String
    PXM_range = c2l(PXM_FIRST_PXM_COL) & PXM_FIRST_DATA_ROW & ":" & _
                c2l(PXM_FIRST_PXM_COL - 1 + num_markers) & (PXM_FIRST_DATA_ROW + num_projects - 1)
    mc_array = Range(PXM_range)
    
    ' copy the marker confidence array in the the crosswalk sheet
    Sheets(EXPERTISE_CROSSWALK_SHEET).Activate
    Dim Destination As Range
    mc_range = c2l(EC_DATA_FIRST_MARKER_COLUMN) & _
                    (EC_DATA_FIRST_MARKER_ROW) & ":" & _
                c2l(EC_DATA_FIRST_MARKER_COLUMN - 1 + num_markers) & _
                    (EC_DATA_FIRST_MARKER_ROW + num_projects - 1)
    Set Destination = Range(mc_range)
    Destination.Resize(UBound(mc_array, 1), UBound(mc_array, 2)).Value = mc_array
        
    ' make a copy of the marker confidence array to help with filling holes in the assignemnts (swapping)
    ReDim mc_as_loaded(1 To num_projects, 1 To num_markers)
    mc_as_loaded = mc_array
    
    ' load the marker number array
    Dim mn_range As String
    mn_range = c2l(EC_DATA_FIRST_MARKER_COLUMN) & _
                    (EC_DATA_FIRST_MARKER_ROW - 1) & ":" & _
                c2l(EC_DATA_FIRST_MARKER_COLUMN - 1 + num_markers) & _
                    (EC_DATA_FIRST_MARKER_ROW - 1)
    mn_array = Range(mn_range)
    
    ' load the counts of ratings (H, M, L, X) per project
    Dim xlmh_projects_range As String
    ' 4 columns wide X num_projects deep
    xlmh_projects_range = c2l(EC_XLMH_CONFIDENCE_PER_PROJECT_COL) & EC_DATA_FIRST_MARKER_ROW & ":" & _
                c2l(EC_XLMH_CONFIDENCE_PER_PROJECT_COL + 3) & (EC_DATA_FIRST_MARKER_ROW + num_projects - 1)
    xlmh_per_project = Range(xlmh_projects_range)
    
    ' array with the confidence of the assignments made (will be updated)
    Dim coa_array_range As String
    ' 'target_markers_per_proj' columns wide (X L M H) by num_projects deep
    coa_array_range = c2l(EC_ASSIGNMENT_CONFIDENCE_FIRST_COLUMN) & EC_DATA_FIRST_MARKER_ROW & ":" & _
                c2l(EC_ASSIGNMENT_CONFIDENCE_FIRST_COLUMN + target_markers_per_proj - 1) & (EC_DATA_FIRST_MARKER_ROW + num_projects - 1)
    coa_array = Range(coa_array_range)
      
    ' load the array of marker total H,M and X's per marker
    Dim xlmh_markers_range As String
    ' num_markers wide X 4 rows deep
    xlmh_markers_range = c2l(EC_DATA_FIRST_MARKER_COLUMN) & EC_XLMH_MARKER_TABLE_FIRST_ROW & ":" & _
                c2l(EC_DATA_FIRST_MARKER_COLUMN + num_markers - 1) & (EC_XLMH_MARKER_TABLE_FIRST_ROW + 3)
    xlmh_per_marker = Range(xlmh_markers_range)
    
    ' Read in the arrays that will be updated with the assignment information.
    ' This allows users to define some assignments, and then let the software complete the assignments
    ' initialize the marker assignment, and confidence of assigned marker arrays (N per project)
    Dim assignments_range As String
    ' target_markers_per_proj wide x num_projects deep
    assignments_range = c2l(EC_ASSIGNMENTS_FIRST_COLUMN) & EC_DATA_FIRST_MARKER_ROW & ":" & _
                c2l(EC_ASSIGNMENTS_FIRST_COLUMN + target_markers_per_proj - 1) & _
                        (EC_DATA_FIRST_MARKER_ROW + num_projects - 1)
    assignments = Range(assignments_range)
            
    ' loop through the projects assigning first those with the least expertise
    Dim i As Long, j As Long, num_assigned As Long
    Dim next_proj As Long        ' # of next project that should be assigned.
    
    ' load the arrays counting numbers assigned from the marker assignments read from the worksheet
    ReDim n_assigned2project(1 To num_projects)
    ReDim n_assigned2marker(1 To num_markers)
    ReDim n_per_assignment_col(1 To target_markers_per_proj)

'DEBUG DEBUG
    ReDim ass_this_col(1 To num_projects)
'END DEBUG
    num_assigned = 0
    For i = 1 To num_projects
        For j = 1 To target_markers_per_proj
            If assignments(i, j) > 0 Then
                n_assigned2project(i) = n_assigned2project(i) + 1
                n_assigned2marker(assignments(i, j)) = n_assigned2marker(assignments(i, j)) + 1
                n_per_assignment_col(j) = n_per_assignment_col(j) + 1
                num_assigned = num_assigned + 1
            End If
        
'DEBUG
            If (j = 1) And (assignments(i, j) > 0) Then
                ass_this_col(i) = 1
            End If
'END DEBUG
        Next j
    Next i
    
    Dim assignment_col As Long          ' assignment column (1 = first reader, 2 = second reader ...
    assignment_col = 1                  ' start with the projects that need the most markers
    Dim best_marker As Long             ' number of marker proposed to review this project
    Dim looking As Boolean
    looking = True
    Dim mentor_num As Long
    Dim num_conflicts As Long
    Dim num_first_reader_assignments() As Long
    ReDim num_first_reader_assignments(1 To num_markers)
    ReDim assignment_failed_for_this_proj(1 To num_projects)
    Dim num_assignments_failed As Long
    Dim last_marker As Long              ' DEBUG DEBUG DEBUG
    
    ' ready to start assigning
    While looking
        'find the project with the lowest available expertise
        next_proj = FindNextProject(assignment_col, next_proj)
        mentor_num = Sheets(PROJECTS_SHEET).Range(c2l(PS_MENTOR_ID_COLUMN) & (PS_FIRST_DATA_ROW + next_proj - 1)).Value
        best_marker = 0
        If (next_proj > 0) Then
            ' for this project, find the marker with the highest available confidence rating
            ' and (if there are multiple possibilities) find the lowest number this confidence ratings available
            For j = 1 To num_markers
                If (mc_array(next_proj, j) = "X") Or _
                   (mc_array(next_proj, j) = "A") Or _
                   (mc_array(next_proj, j) = "") Or _
                   Len(mc_array(next_proj, j)) = 0 Then     ' no data (i.e. no expertise profile)
                    ' don't consider this marker if there is a COI or
                    ' if they have already been assigned to this project or
                    ' they did not provide an expertise sheet
                Else
                    If (n_assigned2marker(j) >= target_ass_per_marker) Or _
                        ((assignment_col = 1) And _
                        (max_first_reader_assignments > 0) And _
                        (num_first_reader_assignments(j) >= max_first_reader_assignments)) Then
                        ' we reached the limit on the number of markers per reader
                        ' or the limit on the number of first reader assignments per reader
                        ' so this marker cannot be considered for this project
                        i = i
                    Else
                        If j <> mentor_num Then
                            ' the candidate marker is not the mentor
                            If best_marker = 0 Then
                                best_marker = j
                            Else
                                If CompareConfidence(next_proj, j, best_marker) Then
                                    best_marker = j
                                End If
                            End If
                        Else
                            MsgBox "Mentor # coding error", vbCritical
                            Exit Function
                        End If
                    End If
                End If
            Next j
            If best_marker > 0 Then
                ' first a couple of checks
                If best_marker = mentor_num Then
                    MsgBox "[AssignMarkersBasedOnConfidence] Proposed marker is also the mentor ?????", vbCritical
                    Exit Function
                End If
                ' make the assignment
                assignments(next_proj, assignment_col) = best_marker
' debug
 last_marker = best_marker
' debug
                ' also store the confidence for this project of the marker assigned
                coa_array(next_proj, assignment_col) = GetConfidenceCode(CDbl(mc_array(next_proj, best_marker)))
                'update the other information arrays since we have removed a project & assigned one to a marker
                UpdateArrays next_proj, best_marker
                
                n_assigned2project(next_proj) = n_assigned2project(next_proj) + 1
                n_assigned2marker(best_marker) = n_assigned2marker(best_marker) + 1
                If assignment_col = 1 Then
                    num_first_reader_assignments(best_marker) = num_first_reader_assignments(best_marker) + 1
                End If
                
' DEBUG
                n_per_assignment_col(assignment_col) = n_per_assignment_col(assignment_col) + 1
                If ass_this_col(next_proj) = 1 Then
'                    MsgBox "should be zero", vbCritical
                End If
                ass_this_col(next_proj) = 1 'DEBUG DEBUG DEBUG flag which projects have been assigned
' END DEBUG
                num_assigned = num_assigned + 1
            Else
                assignment_failed_for_this_proj(next_proj) = True
                num_assignments_failed = num_assignments_failed + 1
'                MsgBox "unable to find available marker for project " & next_proj & ", terminating search", vbCritical
'                looking = False
            End If
        Else
            looking = False ' no more projects need assigning
        End If
    Wend
    
    'see of we can fill in the assignment holes
    If num_assignments_failed > 0 Then
        If FillAssignmentHoles(num_assigned) = False Then
            Exit Function
        End If
    End If
    
    ' update the assignment number (the calculations correspond to marker numbers that go from 1 to num_markers
    ' however since the data is taken from a spreadsheet, the marker columns may have been reordered.
    ' update the assignment numbers to reflect the marker numbers specified in the crosswalk table
    Dim n_empty As Long
    For i = 1 To num_projects
        For j = 1 To target_markers_per_proj
            If IsEmpty(assignments(i, j)) Then
                n_empty = n_empty + 1
            Else
                assignments(i, j) = mn_array(1, assignments(i, j))
            End If
        Next j
    Next i
    If num_assigned = 0 Then
        MsgBox "Nothing assigned - check that inputs have been provided.", vbCritical
        AssignMarkersBasedOnConfidence = False
        Exit Function
    Else
        If n_empty > 0 Then
            MsgBox "NOTE: not all assignments made, " & n_empty & " assignments by hand required!", vbCritical
        End If
    End If
    ' write the marker assignment array
    Set Destination = Range(assignments_range)
    Destination.Resize(UBound(assignments, 1), UBound(assignments, 2)).Value = assignments
    
    ' write out the confidence letter for the selected marker on this project
    Set Destination = Range(coa_array_range)
    Destination.Resize(UBound(coa_array, 1), UBound(coa_array, 2)).Value = coa_array
        
    ' write out the array of marker confidences (updated for assignments)
    Set Destination = Range(mc_range)
    Destination.Resize(UBound(mc_array, 1), UBound(mc_array, 2)).Value = mc_array
    
    ' copy the assignments into the assignment master sheet
    Sheets(MASTER_ASSIGNMENTS_SHEET).Activate
    Dim ass_sht_ass_range As String
    ass_sht_ass_range = c2l(MAS_FIRST_ASSIGNMENT_COLUMN) & MAS_FIRST_ASSIGNMENT_ROW & ":" & _
        c2l(MAS_FIRST_ASSIGNMENT_COLUMN + target_markers_per_proj - 1) & (MAS_FIRST_ASSIGNMENT_ROW + num_projects - 1)
    Set Destination = Range(ass_sht_ass_range)
    Destination.Resize(UBound(assignments, 1), UBound(assignments, 2)).Value = assignments

    Sheets(MASTER_ASSIGNMENTS_SHEET).Activate

    AddMessage "All done assigning markers to projects. Sheet has: " & num_assigned & " marking assignments for " _
            & num_projects & " projects and " & num_markers & " markers."
    Sheets(starting_sheet).Activate
    
    AssignMarkersBasedOnConfidence = True
    
End Function

Function MarkerOnThisProjectAlready(marker_num As Long, project_num As Long) As Boolean
    Dim i As Long
    For i = 1 To n_assigned2project(project_num)
        If assignments(project_num, i) = marker_num Then
            MarkerOnThisProjectAlready = True
            Exit Function
        End If
    Next i
End Function


Public Function FillAssignmentHoles(num_assignments As Long) As Boolean

    ' see if we can make all the remaining assignments by swapping pairs of markers
    ' (one of which does not yet have a full suite of assignments)
    
    ' basic approach - if:
    '   a marker has room for more assignments
    '   and is not in conflict for that project
    '   and the assigned marker is not in conflict for the project needing markers
    '   and marker to insert has confidence on that propsal equal to the existing marker
    ' then:
    '   move the assigned marker to the empty slot
    '   put the marker with room for more assignments into the newly vacated slot

    ' search order:
    '   down the projects in each column of assignments looking for empty slots
    '   down all assigned markers in an assignment slot looking of candidates to pop
    '   through all the markers looking for one that has good enough confidence to replace the marker popped out

    Dim i As Long, j As Long, k As Long, m As Long
    Dim marker2swap As Long, conf2move As Long, conf2insert As Long
    
    For j = 1 To target_markers_per_proj        ' for a given assignment column
        For i = 1 To num_projects               ' for each of the projects assigned readers
            If assignments(i, j) = 0 Then
                ' empty assignment slot is for project 'i' in column 'j'
                ' go through all the project assignments in this assignment_col
                ' looking for one that can be 'popped-out' and used to fill the empty slot
                If num_assignments < num_markers * target_ass_per_marker Then
                    ' there are still markers that could be given an additional marking assignment
                    For k = 1 To num_projects
                        'look through the assignments in this column for markers to be popped out
                        marker2swap = assignments(k, j)
                        If (k <> i) And marker2swap > 0 Then
                            If (MarkerOnThisProjectAlready(marker2swap, i) = False) And _
                                (mc_as_loaded(k, marker2swap) <> "A") And _
                                (mc_as_loaded(i, marker2swap) <> "X") Then
                                ' for each marker assignment slot in this column
                                    ' exclude if:
                                        ' this is the marker we are looking at ejecting
                                        ' there is not a marker assigned (another empty slot)
                                        ' the marker was not assigned to this project before this assignment run (i.e., fixed)
                                        ' they don't have a conflict with this project
                                ' marker2swap looks like a candidate to be popped, and replaced by
                                ' a marker who does not have a full of assignments yet
                                For m = 1 To num_markers
                                    ' basic approach to find a marker to replace the one popped: if:
                                    '   a marker has room for more assignments
                                    '   has at least the same confidence code for a project as the marker assigned
                                    '   and is not in conflict for that project
                                    If (n_assigned2marker(m) < target_ass_per_marker) And _
                                        (marker2swap <> m) And _
                                        (MarkerOnThisProjectAlready(m, i) = False) And _
                                        (mc_as_loaded(k, m) <> "X") Then
                                        ' compare the confidences about project k for
                                        ' this marker (m) and marker2swap (...2move)
                                        conf2move = GetConfidenceCode(CDbl(mc_as_loaded(k, assignments(k, j))))
                                        conf2insert = GetConfidenceCode(CDbl(mc_as_loaded(k, m)))
                                        If conf2insert >= conf2move Then
                                            ' OK to replace marker2swap with m
                                            assignments(i, j) = marker2swap
                                            assignments(k, j) = m
                                            num_assignments = num_assignments + 1
                                            m = num_markers ' exit this loop
                                            k = num_projects
                                            n_assigned2marker(m) = n_assigned2marker(m) + 1
                                        End If
                                    End If
                                Next m
                            End If
                        End If
                    Next k
                Else
                    ' no more markers should get assignments so we are done
                    FillAssignmentHoles = True
                    Exit Function
                End If
            End If
        Next i
    Next j
    
    FillAssignmentHoles = True
End Function

Public Function FindNextProject(assignment_col As Long, last_proj_assigned As Long) As Long
' from among the projects that need  more reader/markers at the current 'assignment_col'
' How? find the project with the lowest available confidence,
' particularly the fewest high-confidence reviewers.
' if no more projects to assign at this level, move to the next level

    ' assignment_col is the column of the assignment array currently being filled with marker #'s)
    Dim next_project2assign As Long, this_one_is_better As Boolean
    Dim num_highs As Long, num_meds As Long, num_lows As Long
    ' initialize
    num_highs = 2 * num_projects
    num_meds = 2 * num_projects
    num_lows = 2 * num_projects
    Dim looking As Boolean
    looking = True
    Dim i As Long
    i = 0    'go through the full project list, as the next project to assign could now be before the last one assigned
    While looking
        i = i + 1
        If (i <= num_projects) Then
            ' run through all the projects
            If (assignment_failed_for_this_proj(i) = False) And (assignments(i, assignment_col) = 0) Then
                ' still worth trying, and this project# needs a marker in this assignment_col
                If xlmh_per_project(i, 4) < num_highs Then
                    ' this project has fewer people rating this as @High confidence (compared to current choice)
                    this_one_is_better = True
                Else
                    If xlmh_per_project(i, 4) = num_highs Then
                        If xlmh_per_project(i, 3) < num_meds Then
                            ' they are the same @High, but this one has fewer @Medium's
                            this_one_is_better = True
                        Else
                            If xlmh_per_project(i, 3) = num_meds Then
                                ' They are the same @High and @Medium but this one has fewer @Lows
                                If xlmh_per_project(i, 2) < num_lows Then
                                    this_one_is_better = True
                                Else
                                    ' the current 'next_project2assign is still a better choice
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If this_one_is_better Then
            num_highs = xlmh_per_project(i, 4)
            num_meds = xlmh_per_project(i, 3)
            num_lows = xlmh_per_project(i, 2)
            next_project2assign = i
            this_one_is_better = False
        End If
        If i = num_projects Then
            ' we have gone through all the projects looking for assignment candidates in this column
            If (num_highs > num_projects) And (num_meds > num_projects) Then
                ' no projects found needing markers at this level
'                If n_per_assignment_col(assignment_col) <> num_projects Then
'                    MsgBox n_per_assignment_col(assignment_col) & " of " & _
'                    num_projects & " had full assignments, manual assigning required!", vbOKOnly
'                End If
                If assignment_col = target_markers_per_proj Then
                    ' no more levels, we are done looking for projects to assign
                    FindNextProject = 0   ' flag that there are no projects that need a reader assigned
                    Exit Function
                Else
                    ' move to the next level
                    assignment_col = assignment_col + 1
                    i = 0
                End If
            Else
                ' a candidate was found, return it
                FindNextProject = next_project2assign
                looking = False
            End If
        End If
    Wend

End Function

Function UpdateArrays(project_num As Long, marker_num As Long) As Boolean
    ' since a marker has been assigned to a project, the number of markers available for the project
    ' and the number of available confidence specifications for a marker has been reduced
    Dim conf As String
    conf = GetConfidenceCode(CDbl(mc_array(project_num, marker_num)))
    Select Case conf
    Case 3
        xlmh_per_project(project_num, 4) = xlmh_per_project(project_num, 4) - 1
        xlmh_per_marker(4, marker_num) = xlmh_per_marker(4, marker_num) - 1
    Case 2
        xlmh_per_project(project_num, 3) = xlmh_per_project(project_num, 3) - 1
        xlmh_per_marker(3, marker_num) = xlmh_per_marker(3, marker_num) - 1
    Case 1
        xlmh_per_project(project_num, 2) = xlmh_per_project(project_num, 2) - 1
        xlmh_per_marker(2, marker_num) = xlmh_per_marker(2, marker_num) - 1
    Case Else
        MsgBox "Error: marker selected for a project they are in conflict for", vbCritical
        UpdateArrays = False
        Exit Function
    End Select
    mc_array(project_num, marker_num) = "A"   'flag the marker has been assigned to this project
    UpdateArrays = True
    
End Function

Public Function CompareConfidence(this_proj As Long, this_marker As Long, best_marker As Long) As Boolean
    ' select between two candidate markers. Choose the one with the higher confidence.
    ' in case of a tie choose the one with more confidence rankings at this level
    
    Dim best_ranking As Long, this_ranking As Long
    Dim this_marker_num_available As Long, best_marker_num_available As Long
    
    Dim best_letter As String, this_letter As String
    If mc_array(this_proj, this_marker) > mc_array(this_proj, best_marker) Then
        best_marker = this_marker
        CompareConfidence = True
    Else
        best_ranking = GetConfidenceCode(CDbl(mc_array(this_proj, best_marker)))
        this_ranking = GetConfidenceCode(CDbl(mc_array(this_proj, this_marker)))
        If this_ranking = best_ranking Then
            this_marker_num_available = xlmh_per_marker(this_ranking + 1, this_marker)
            best_marker_num_available = xlmh_per_marker(this_ranking + 1, best_marker)
            If this_marker_num_available > best_marker_num_available Then
                best_marker = this_marker
                CompareConfidence = True
            Else
                CompareConfidence = False
            End If
        End If
    End If
End Function


Public Function GetConfidenceCode(confidence_level As Double) As Long
    If confidence_level < 0 Then
        MsgBox "confidence level = " & confidence_level & " should be 0 to 1", vbCritical
        GetConfidenceCode = -1
    Else
        If confidence_level <= ONE_THIRD Then
            GetConfidenceCode = 1
        Else
            If confidence_level <= TWO_THIRDS Then
                GetConfidenceCode = 2
            Else
                If confidence_level <= 1 Then
                    GetConfidenceCode = 3
                Else
                    MsgBox "confidence level = " & confidence_level & " should be 0 to 1", vbCritical
                    GetConfidenceCode = -1
                End If
            End If
        End If
    End If
End Function

Public Function ClearMss() As Boolean
'
' return the Master Scoresheet to a state that removes any scores added.
    Sheets(MASTER_SCORESHEET).Activate
    
    ' build the expression for the columns of scores to clear
    Dim fr As Long, lr As Long, fc As Long
    fr = MSS_FIRST_PROJECT_ROW
    fc = MSS_FIRST_SCORE_COL
    lr = fr + max_projects - 1
    Dim clear_range As String, i As Long
    For i = 1 To max_markers_per_proj
        If i > 1 Then
            clear_range = clear_range & ","
        End If
        clear_range = clear_range & c2l(fc) & fr & ":" & c2l(fc + max_criteria - 1) & lr
        fc = fc + max_criteria + 2
    Next i
    
    '    "H6:M1005,o6:T1005,V6:AA1005,AC6:AH1005,AJ6:Ao1005"

    Dim start_address As String, focus_cell As String
    start_address = ActiveCell.Address
    focus_cell = FirstCell(clear_range)
    Range(focus_cell).Select
    Range(clear_range).Select
    Range(focus_cell).Activate
    Selection.ClearContents
    Range(focus_cell).Select
    ClearMss = True
    
'    Range(start_address).Select
End Function

Function AddMessage(msg2add As String) As Long
    If buffer_messages Then
        num_messages = num_messages + 1
        ReDim Preserve messages(1 To num_messages)
        messages(num_messages) = msg2add
        AddMessage = num_messages
    Else
        MsgBox msg2add, vbOKOnly
    End If
End Function

Function InitMessages()
    buffer_messages = True
    num_messages = 0
End Function

Function ReportMessages()
    Dim msg_out As String
    If num_messages > 0 Then
        Dim i As Long
        For i = 1 To num_messages
            msg_out = msg_out & messages(i)
            If i <> num_messages Then
                msg_out = msg_out & vbCrLf
            End If
        Next i
        MsgBox msg_out, vbOKOnly
    End If
    
End Function

Function clear_crosswalk_sheet(crosswalk_sheet As String) As Boolean
    Const EC_DATA_RANGE As String = "J7:SY1006"
    With ActiveWorkbook.Sheets(crosswalk_sheet)
        .Select
        .Range("J7:SY1006").Select
    End With
    Selection.ClearContents
    Range(FirstCell(EC_DATA_RANGE)).Select

End Function

Sub ExportCompetitionWorkbookSub()
   
   ExportCompetitionWorkbook

End Sub

Public Function ClearKeywordTablesPXMCrosswalkAssignmentsAndMss()

    If DefineGlobals = False Then
        ClearKeywordTablesPXMCrosswalkAssignmentsAndMss = False
        Exit Function
    End If
    
    Dim start_sheet As String
    start_sheet = ActiveSheet.Name
    
    ClearKeywordTables
    
    ' PXM sheet
    ClearPXMSheet
    
    ' crosswalk sheet
    Sheets(EXPERTISE_CROSSWALK_SHEET).Activate
    Sheets(EXPERTISE_CROSSWALK_SHEET).Select
    clear_crosswalk_sheet ActiveSheet.Name
    
    'assignment sheet
    ClearAssignmentSheet
        
    ' master scoresheet
    ClearMss                'clear the sheet in case it already has scores in it.
    UnsortMasterScoresheet  ' make sure the projects go from 1 to N
    
    Dim i As Long
    For i = 1 To Sheets.Count
        If Sheets(i).Name = SHARED_SCORESHEET Then
            Application.DisplayAlerts = False ' don't ask for confirmation
            Sheets(i).Delete
            Application.DisplayAlerts = True
            ' we're done, so clean up and exit the function
            Sheets(start_sheet).Select
            Exit Function
        End If
    Next i
    ' shared scoresheet not found, return to the starting sheet
    Sheets(start_sheet).Select
    
End Function
 
Function RemoveReferencesToOtherWorkbooks() As Boolean
    ' replace the contents of cells in the competition book with formula references to the macro book
    ' with the current text contents (since the text should stay static from this point)
        
    'three 'max' parameters from the systems parameters sheet
    formulas2text COMPETITION_PARAMETERS_SHEET, "D2:D5"
    formulas2text COMPETITION_PARAMETERS_SHEET, "C11"
    formulas2text PROJECT_KEYWORDS_SHEET, "A1:" & c2l(2 * max_keywords + 6) & 2
    formulas2text MARKER_EXPERTISE_SHEET, "B2:" & c2l(2 * max_keywords + 6) & 2
    formulas2text MASTER_ASSIGNMENTS_SHEET, MAS_FLAG_COLUMN & "2:" & MAS_FLAG_COLUMN & "3"
    
    RemoveReferencesToOtherWorkbooks = True
End Function

Public Function formulas2text(sheet_name As String, range2change As String) As Boolean
    
    Dim current_cell As String, current_sheet As String, editsheet_celladdress As String
    current_sheet = ActiveSheet.Name
    current_cell = ActiveCell.Address
    Sheets(sheet_name).Select
    editsheet_celladdress = ActiveCell.Address
    Range(range2change).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range(editsheet_celladdress).Activate
    Sheets(current_sheet).Select
    Range(current_cell).Activate
    
    formulas2text = True
End Function

Function ExportCompetitionWorkbook() As Boolean
'
    ' choose the sheets, and copy them into a new workbook
    On Error GoTo Shts_missing
    Sheets(Array(COMPETITION_PARAMETERS_SHEET, CRITERIA_SHEET, PROJECTS_SHEET, MARKERS_SHEET, KEYWORDS_SHEET, _
        PROJECT_KEYWORDS_SHEET, MARKER_EXPERTISE_SHEET, PROJECT_X_MARKER_SHEET, _
        EXPERTISE_CROSSWALK_SHEET, MASTER_ASSIGNMENTS_SHEET, MASTER_SCORESHEET, SHARED_SCORESHEET)).Select
    If False Then
Shts_missing:
        MsgBox "Error selecting the sheets to export, are they all there?", vbCritical
        ExportCompetitionWorkbook = False
        On Error GoTo 0
        On Error Resume Next
        Exit Function
    Else
        Sheets(Array(COMPETITION_PARAMETERS_SHEET, CRITERIA_SHEET, PROJECTS_SHEET, MARKERS_SHEET, KEYWORDS_SHEET, _
            PROJECT_KEYWORDS_SHEET, MARKER_EXPERTISE_SHEET, PROJECT_X_MARKER_SHEET, _
            EXPERTISE_CROSSWALK_SHEET, MASTER_ASSIGNMENTS_SHEET, MASTER_SCORESHEET, SHARED_SCORESHEET)).Select
    End If
    On Error Resume Next
        
    Sheets(MASTER_SCORESHEET).Activate
    Sheets(Array(COMPETITION_PARAMETERS_SHEET, CRITERIA_SHEET, PROJECTS_SHEET, MARKERS_SHEET, KEYWORDS_SHEET, _
        PROJECT_KEYWORDS_SHEET, MARKER_EXPERTISE_SHEET, PROJECT_X_MARKER_SHEET, _
        EXPERTISE_CROSSWALK_SHEET, MASTER_ASSIGNMENTS_SHEET, MASTER_SCORESHEET, SHARED_SCORESHEET)).Copy
    Dim new_workbook As String
    new_workbook = ActiveWorkbook.Name
    
    ' clear the worksheets in the workbook with the macros (to be ready to start again)
    ThisWorkbook.Activate
    
    ClearKeywordTablesPXMCrosswalkAssignmentsAndMss    ' this also remove the copy of the shared scoresheet
    
    ' go back and ask the user to save the new workbook
    Workbooks(new_workbook).Activate

    ' remove references to the macro workbook
    If RemoveReferencesToOtherWorkbooks = False Then
        Exit Function
    End If
    
    '  and clear the master scoresheet in the new book
    Sheets(MASTER_SCORESHEET).Activate
    ClearMss   'Make sure the sheet is ready to recieve user scores
    UnsortMasterScoresheet  ' make sure the projects go from 1 to N
    
    Sheets(MASTER_ASSIGNMENTS_SHEET).Select
    ' adjust the slider so that more tabs show
    ActiveWindow.TabRatio = 0.855
    
    ' get the saveas name
    Dim fileSaveName As String
    fileSaveName = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx")
    If Len(fileSaveName) = 0 Then
        MsgBox "File not saved!", vbOKOnly
    End If
    
    'save it
    ActiveWorkbook.SaveAs filename:=fileSaveName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        
    ExportCompetitionWorkbook = True
    
End Function

Public Function ClearAssignmentSheet() As Boolean

    Sheets(MASTER_ASSIGNMENTS_SHEET).Activate
    Dim ass_sht_ass_range As String
    ass_sht_ass_range = c2l(MAS_FIRST_ASSIGNMENT_COLUMN) & MAS_FIRST_ASSIGNMENT_ROW & ":" & _
        c2l(MAS_FIRST_ASSIGNMENT_COLUMN + target_markers_per_proj - 1) & (MAS_FIRST_ASSIGNMENT_ROW + num_projects - 1)
    Range(ass_sht_ass_range).Select
    Selection.ClearContents
    ClearAssignmentSheet = True
    
End Function
Public Function CreateFigureOfMeritTable() As Boolean
    
    If (CreatePXMFromProjectRelevanceAndMarkerExpertise = False) Then
        Exit Function
    End If
    CreateFigureOfMeritTable = True
End Function

Public Function is_array_empty(array2test() As Variant) As Boolean
    Dim i As Long, j As Long
    Dim upper1 As Long, upper2 As Long
    upper1 = UBound(array2test, 1)
    upper2 = UBound(array2test, 2)
    For i = 1 To upper1
        For j = 1 To upper2
            If Len(array2test(i, j)) > 0 Then
                is_array_empty = False
                Exit Function
            End If
        Next j
    Next i
    
    is_array_empty = True
    
End Function
Public Function CreatePXMFromProjectRelevanceAndMarkerExpertise() As Boolean
' calculate the Figure of Merits (FOM) which estimate the confidence of a marker to review a project
' given:
' a table of marker confidence with regards to a number of keywords or themes
' a table of the relevance of projects to the same keywords or themes
' weights on each of the keywords
' then the FOM for marker i's confidence to review project j is estimated as the sum of the products of
'   (the marker's confidence for a keyword times
'   the project's relevance to that keyword
'   the weight of that keyword)
'   divided by the sum of the weights

    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    Dim i As Long, j As Long, k As Long
    Dim kw_range As String, me_range As String, pk_range As String, PXM_range As String
    Dim kw_weights() As Variant, me_array() As Variant, pk_array() As Variant
    Dim pxm() As Variant
    ReDim pxm(1 To num_projects, 1 To num_markers)
    
    ' read in the keyword weights
    Sheets(KEYWORDS_SHEET).Activate
    kw_range = c2l(KW_WEIGHTS_COL) & KW_WEIGHTS_ROW & ":" & _
               c2l(KW_WEIGHTS_COL) & (KW_WEIGHTS_ROW - 1 + num_keywords)
    kw_weights = Range(kw_range)
    If is_array_empty(kw_weights) = True Then
        AddMessage "Keyword weights array is empty, check table in " & KEYWORDS_SHEET & " sheet."
        CreatePXMFromProjectRelevanceAndMarkerExpertise = False
        Exit Function
    End If
    
    ' read in the marker relevances to keywords
    Sheets(MARKER_EXPERTISE_SHEET).Activate
    me_range = c2l(ME_FIRST_MARKER_DATA_COL) & ME_FIRST_MARKER_DATA_ROW & ":" & _
               c2l(ME_FIRST_MARKER_DATA_COL - 1 + num_keywords) & (ME_FIRST_MARKER_DATA_ROW + num_markers - 1)
    me_array = Range(me_range)
    If is_array_empty(me_array) = True Then
        AddMessage "Array of marker confidence on keywords is empty, check table in " & _
                    MARKER_EXPERTISE_SHEET & " sheet."
        CreatePXMFromProjectRelevanceAndMarkerExpertise = False
        Exit Function
    End If

    ' read in the project keyword confidences
    Sheets(PROJECT_KEYWORDS_SHEET).Activate
    pk_range = c2l(PK_FIRST_PROJECT_DATA_COL) & PK_FIRST_PROJECT_DATA_ROW & ":" & _
               c2l(PK_FIRST_PROJECT_DATA_COL - 1 + num_keywords) & (PK_FIRST_PROJECT_DATA_ROW + num_projects - 1)
    pk_array = Range(pk_range)
    If is_array_empty(pk_array) = True Then
        AddMessage "Array of project ratings by keyword is empty, check table in " & _
                    PROJECT_KEYWORDS_SHEET & " sheet."
        CreatePXMFromProjectRelevanceAndMarkerExpertise = False
        Exit Function
    End If

    ' read in the column of mentor numbers for projects
    Sheets(PROJECTS_SHEET).Activate
    Dim mentor_range As String
    mentor_range = (c2l(PS_MENTOR_ID_COLUMN) & PS_FIRST_DATA_ROW) & ":" & _
                   (c2l(PS_MENTOR_ID_COLUMN) & PS_FIRST_DATA_ROW + num_projects - 1)
    mentor_column = Range(mentor_range)
    
    ' read in the PXM table to get the conflicts of interest
    Sheets(PROJECT_X_MARKER_SHEET).Activate
    PXM_range = c2l(PXM_FIRST_PXM_COL) & PXM_FIRST_DATA_ROW & ":" & _
                c2l(PXM_FIRST_PXM_COL - 1 + num_markers) & (PXM_FIRST_DATA_ROW + num_projects - 1)
    pxm = Range(PXM_range)
    
    ' calculate the PXM ratings
    Dim max As Double, num_conflicts As Long
    max = 0
    num_conflicts = 0
    For i = 1 To num_projects
        For j = 1 To num_markers
            If pxm(i, j) <> "X" Then
                pxm(i, j) = 0
                For k = 1 To num_keywords
                    pxm(i, j) = pxm(i, j) + _
                    kw_weights(k, 1) * LMH2Percent(CStr(pk_array(i, k))) * LMH2Percent(CStr(me_array(j, k)))
                Next k
                If max < pxm(i, j) Then
                    max = pxm(i, j)
                End If
            Else
                num_conflicts = num_conflicts + 1
            End If
        Next j
    Next i
    
    ' scale the PXM to go from zero to one
    For i = 1 To num_projects
        For j = 1 To num_markers
            If pxm(i, j) <> "X" Then
                If mentor_column(i, 1) <> j Then
                    pxm(i, j) = pxm(i, j) / max
                Else
                    pxm(i, j) = "X" 'flag the mentor is in conflict for this project
                    num_conflicts = num_conflicts + 1
                End If
            End If
        Next j
    Next i
    
    ' write the PXM array
    Dim Destination As Range
    Set Destination = Range(PXM_range)
    Destination.Resize(UBound(pxm, 1), UBound(pxm, 2)).Value = pxm
    Range(FirstCell(PXM_range)).Activate
    
    CreatePXMFromProjectRelevanceAndMarkerExpertise = True
    
End Function

Public Function ClearPXMSheet() As Boolean
    
    Sheets(PROJECT_X_MARKER_SHEET).Activate
    Dim PXM_range As String
    PXM_range = c2l(PXM_FIRST_PXM_COL) & PXM_FIRST_DATA_ROW & ":" & _
                c2l(PXM_FIRST_PXM_COL - 1 + max_markers) & (PXM_FIRST_DATA_ROW + max_projects - 1)
    Range(PXM_range).Select
    Range(FirstCell(PXM_range)).Activate
    Selection.ClearContents
    Range(FirstCell(PXM_range)).Select
    Range(FirstCell(PXM_range)).Activate
    ActiveWindow.Zoom = 100

    ClearPXMSheet = True
    
End Function

Public Function FindHeaderColumn(row_num As Long, search_text As String, exact As Boolean) As Long
    ' find and move to a cell in the specified row that contains the search text
    Dim found As Boolean
    Dim search_flag As Long
    If exact = True Then
        search_flag = xlWhole
    Else
        search_flag = xlPart
    End If
    Rows(row_num).Select
    
    If Selection.Find(What:=search_text, after:=ActiveCell, LookIn:=xlFormulas2, _
        LookAt:=search_flag, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False) Is Nothing Then
        'nothing found, nothing for this marker so exit the loop
        found = False
    Else
        Selection.Find(What:=search_text, after:=ActiveCell, LookIn:=xlFormulas2, _
            LookAt:=search_flag, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False).Activate
        found = True
    End If
    If found Then
        FindHeaderColumn = ActiveCell.Column
    Else
        FindHeaderColumn = 0
    End If
End Function

Public Function ClearKeywordTables() As Boolean

    Dim start_sheet As String
    start_sheet = ActiveSheet.Name
    
    Dim clear_range As String
    Sheets(PROJECT_KEYWORDS_SHEET).Select
    clear_range = c2l(PK_FIRST_PROJECT_DATA_COL) & PK_FIRST_PROJECT_DATA_ROW & ":" & _
    c2l(PK_FIRST_PROJECT_DATA_COL + max_keywords - 1) & (PK_FIRST_PROJECT_DATA_ROW + max_markers - 1)
    Range(clear_range).Clear
    Range(FirstCell(clear_range)).Select
    
    Sheets(MARKER_EXPERTISE_SHEET).Select
    clear_range = c2l(ME_FIRST_MARKER_DATA_COL) & ME_FIRST_MARKER_DATA_ROW & ":" & _
    c2l(ME_FIRST_MARKER_DATA_COL + max_keywords - 1) & (ME_FIRST_MARKER_DATA_ROW + max_markers - 1)
    Range(clear_range).Clear
    Range(FirstCell(clear_range)).Select
  
    Sheets(start_sheet).Select
    
    ClearKeywordTables = True
End Function

Public Function PopulateSharedScoresheet() As Boolean

    globals_defined = False
    If DefineGlobals = False Then
        Exit Function
    End If
    
    Dim start_sheet As String
    start_sheet = ActiveSheet.Name
    
    ' remove the current shared scoresheet from macro book if it is still hanging around.
    Dim i As Long
    For i = 1 To Sheets.Count
        If Sheets(i).Name = SHARED_SCORESHEET Then
            Application.DisplayAlerts = False ' don't ask for confirmation
            Sheets(i).Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next i
    ' create the shared scoresheet as a copy of the template (since we will convert formulae to text)
    Sheets(SHARED_SCORESHEET_TEMPLATE).Select
    Sheets(SHARED_SCORESHEET_TEMPLATE).Copy after:=Sheets(MASTER_SCORESHEET)
    ActiveSheet.Name = SHARED_SCORESHEET
    Sheets(SHARED_SCORESHEET).Select
    ' set the tab colour to light green
    With ActiveWorkbook.Sheets(SHARED_SCORESHEET).Tab
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.8
    End With

    ' load the assignment information
    Sheets(MASTER_ASSIGNMENTS_SHEET).Select
    Dim assignments_range As String
    assignments_range = c2l(MAS_FIRST_ASSIGNMENT_COLUMN) & MAS_FIRST_ASSIGNMENT_ROW & ":" & _
    c2l(MAS_FIRST_ASSIGNMENT_COLUMN + target_markers_per_proj - 1) & MAS_FIRST_ASSIGNMENT_ROW + num_projects - 1
    assignments = Range(assignments_range)
        
    'populate the left hand table with:
    '   for each marker, one row for each of their marking assignments,
    '   including the corresponding project #
    Dim marker_nums() As Variant
    ReDim marker_nums(1 To num_projects * target_markers_per_proj, 1 To 1)
    ReDim project_nums(1 To num_projects * target_markers_per_proj, 1 To 1)
    Dim row_num As Long, j As Long
    row_num = 0
    For i = 1 To num_projects
        For j = 1 To target_markers_per_proj
            If Len(assignments(i, j)) > 0 Then
                row_num = row_num + 1
                project_nums(row_num, 1) = i
                marker_nums(row_num, 1) = assignments(i, j)
            End If
        Next j
    Next i
    
    If (row_num < 2) Then
        MsgBox "[PopulateSharedScoresheet] Not enough assignments to have a competition", vbCritical
        Exit Function
    End If
    
    'put these arrays in the shared sheet
    Sheets(SHARED_SCORESHEET).Select
    Dim project_nums_range As String, marker_nums_range As String
    project_nums_range = c2l(SS_PROJECT_NUM_COLUMN) & SS_FIRST_DATA_ROW & ":" & _
        c2l(SS_PROJECT_NUM_COLUMN) & (SS_FIRST_DATA_ROW + UBound(project_nums, 1) - 1)
    Dim Destination As Range
    Set Destination = Range(project_nums_range)
    Destination.Resize(UBound(project_nums, 1)).Value = project_nums

    marker_nums_range = c2l(SS_MARKER_NUM_COLUMN) & SS_FIRST_DATA_ROW & ":" & _
        c2l(SS_MARKER_NUM_COLUMN) & (SS_FIRST_DATA_ROW + UBound(project_nums, 1) - 1)
    Set Destination = Range(marker_nums_range)
    Destination.Resize(UBound(marker_nums, 1)).Value = marker_nums

    ' sort the data by increasing marker, and then ascending project numbers
    Dim sort_range As String, key1_range As String, key2_range As String
    sort_range = c2l(SS_FIRST_SORT_COLUMN) & SS_FIRST_DATA_ROW & ":" & _
                c2l(SS_LAST_SORT_COLUMN) & (SS_FIRST_DATA_ROW + row_num + 1)
    key1_range = "A" & SS_FIRST_DATA_ROW & ":" & "A" & (SS_FIRST_DATA_ROW + row_num + 1) 'first sort by marker
    key2_range = "C" & SS_FIRST_DATA_ROW & ":" & "C" & (SS_FIRST_DATA_ROW + row_num + 1) ' then sort by project
    With ActiveWorkbook.Worksheets(SHARED_SCORESHEET).Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=Range(key1_range), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=Range(key2_range), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange Range(sort_range)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
            
    'convert the various lookup formulas in some columns to their text result:
    ConvertCellsDownFromFormula2Text SS_FIRST_DATA_ROW, "B"     '   marker name
    ConvertCellsDownFromFormula2Text SS_FIRST_DATA_ROW, "D"     '   project name
    ConvertCellsDownFromFormula2Text SS_FIRST_DATA_ROW, "T"     '   Marker name
    ConvertCellsDownFromFormula2Text SS_FIRST_DATA_ROW, "AQ"     '   project name
    
    ' the criteria names and scoring ranges also need to be converted to text
    ConvertRangeToText "F2:" & c2l(6 + max_criteria - 1) & "4"
       
    PopulateSharedScoresheet = True
    
End Function

Public Function HideUnusedColumnsRowsF() As Boolean

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim hide_range As String
    If (DefineGlobals = False) Then
        Exit Function
    End If
        
    UnHideAllRowsColumns PROJECT_KEYWORDS_SHEET
    hide_range = c2l(PK_FIRST_PROJECT_DATA_COL + num_keywords) & ":" & _
                 c2l(PK_FIRST_PROJECT_DATA_COL + max_keywords - 1)
    hide_range = hide_range & "," & c2l(PK_FIRST_NORMALIZED_DATA_COL + num_keywords) & ":" & _
                 c2l(PK_FIRST_NORMALIZED_DATA_COL + max_keywords - 1)
    Range(hide_range).Select
    Selection.EntireColumn.Hidden = True
    hide_range = (PK_FIRST_PROJECT_DATA_ROW + num_projects) & ":" & _
                 (PK_FIRST_PROJECT_DATA_ROW + max_projects)
    Range(hide_range).Select
    Selection.EntireRow.Hidden = True
    Range(c2l(PK_FIRST_PROJECT_DATA_COL + num_keywords) & PK_FIRST_PROJECT_DATA_ROW).Activate
        
    UnHideAllRowsColumns MARKER_EXPERTISE_SHEET
    hide_range = c2l(ME_FIRST_MARKER_DATA_COL + num_keywords) & ":" & _
                 c2l(ME_FIRST_MARKER_DATA_COL + max_keywords - 1)
    hide_range = hide_range & "," & c2l(ME_FIRST_NORMALIZED_DATA_COL + num_keywords) & ":" & _
                 c2l(ME_FIRST_NORMALIZED_DATA_COL + max_keywords - 1)
    Range(hide_range).Select
    Selection.EntireColumn.Hidden = True
    hide_range = (ME_FIRST_MARKER_DATA_ROW + num_markers) & ":" & _
                 (ME_FIRST_MARKER_DATA_ROW + max_markers)
    Range(hide_range).Select
    Selection.EntireRow.Hidden = True
    Range(c2l(ME_FIRST_MARKER_DATA_COL) & ME_FIRST_MARKER_DATA_ROW).Activate
    
    UnHideAllRowsColumns PROJECT_X_MARKER_SHEET
    hide_range = c2l(PXM_FIRST_PXM_COL + num_markers) & ":" & _
                 c2l(PXM_FIRST_PXM_COL + max_markers - 1)
    Columns(hide_range).Select
    Selection.EntireColumn.Hidden = True
    hide_range = (PXM_FIRST_PXM_COL + num_projects) & ":" & _
                 (PXM_FIRST_PXM_COL + num_projects + max_projects - 1)
    Rows(hide_range).Select
    Selection.EntireRow.Hidden = True
    Range(c2l(PXM_FIRST_PXM_COL) & PXM_FIRST_DATA_ROW).Activate
        
    UnHideAllRowsColumns EXPERTISE_CROSSWALK_SHEET
    hide_range = c2l(EC_ASSIGNMENT_CONFIDENCE_FIRST_COLUMN + target_markers_per_proj) & ":" & _
                 c2l(EC_ASSIGNMENT_CONFIDENCE_FIRST_COLUMN + max_markers_per_proj - 1)
    hide_range = hide_range & "," & c2l(EC_ASSIGNMENTS_FIRST_COLUMN + target_markers_per_proj) & ":" & _
                 c2l(EC_ASSIGNMENTS_FIRST_COLUMN + max_markers_per_proj - 1)
    Range(hide_range).Select
    Selection.EntireColumn.Hidden = True
    hide_range = c2l(EC_DATA_FIRST_MARKER_COLUMN + num_markers) & ":" & _
                 c2l(EC_DATA_FIRST_MARKER_COLUMN + max_markers - 1)
    Range(hide_range).Select
    Selection.EntireColumn.Hidden = True
    hide_range = (EC_DATA_FIRST_MARKER_ROW + num_projects) & ":" & _
                 (EC_DATA_FIRST_MARKER_ROW + max_projects)
    Range(hide_range).Select
    Selection.EntireRow.Hidden = True
    Range(c2l(EC_ASSIGNMENT_CONFIDENCE_FIRST_COLUMN) & EC_DATA_FIRST_MARKER_ROW).Activate
    
    UnHideAllRowsColumns MASTER_ASSIGNMENTS_SHEET
    hide_range = c2l(MAS_FIRST_ASSIGNMENT_COLUMN + target_markers_per_proj) & ":" & _
                 c2l(MAS_FIRST_ASSIGNMENT_COLUMN + max_markers_per_proj - 1)
    hide_range = hide_range & "," & _
                c2l(MAS_FIRST_ASSIGNMENT_COLUMN + max_markers_per_proj + target_markers_per_proj) & ":" & _
                c2l(MAS_FIRST_ASSIGNMENT_COLUMN + 2 * max_markers_per_proj)
    Range(hide_range).Select
    Selection.EntireColumn.Hidden = True
    hide_range = (MAS_FIRST_ASSIGNMENT_ROW + num_projects) & ":" & _
                 (MAS_FIRST_ASSIGNMENT_ROW + max_projects)
    Range(hide_range).Select
' THIS DOES NOT WORK SEE
' https://support.microsoft.com/en-gb/office/why-do-i-see-a-cannot-shift-objects-off-sheet-message-in-excel-559f37da-2b7f-4548-a58d-96669f5310d6?ui=en-us&rs=en-gb&ad=gb
'
   Selection.EntireRow.Hidden = True
    Range(c2l(MAS_FIRST_ASSIGNMENT_COLUMN) & MAS_FIRST_ASSIGNMENT_ROW).Activate
    
    UnHideAllRowsColumns MASTER_SCORESHEET
    ' first hide the columns for unused criteria
    Dim i As Long
    hide_range = ""
    For i = 1 To max_markers_per_proj
        If i > 1 Then
            hide_range = hide_range & "," & _
                c2l(MSS_FIRST_SCORE_COL + (i - 1) * (max_criteria + 2) + num_criteria) & ":" & _
                c2l(MSS_FIRST_SCORE_COL + (i - 1) * (max_criteria + 2) + max_criteria - 1)
        Else
            hide_range = _
                c2l(MSS_FIRST_SCORE_COL + (i - 1) * (max_criteria + 2) + num_criteria) & ":" & _
                c2l(MSS_FIRST_SCORE_COL + (i - 1) * (max_criteria + 2) + max_criteria - 1)
        End If
    Next i
    Range(hide_range).Select
    Selection.EntireColumn.Hidden = True
    Range(c2l(MSS_FIRST_SCORE_COL) & MSS_FIRST_PROJECT_ROW).Activate
    ' hide the columns for unused readers
    hide_range = c2l(MSS_FIRST_SCORE_COL + (max_criteria + 2) * target_markers_per_proj - 1) & ":" & _
                 c2l(MSS_FIRST_SCORE_COL + (max_criteria + 2) * max_markers_per_proj - 2)
    Range(hide_range).Select
    Selection.EntireColumn.Hidden = True
    ' hide the unused project rows
    hide_range = (MSS_FIRST_PROJECT_ROW + num_projects) & ":" & _
                 (MSS_FIRST_PROJECT_ROW + max_projects)
    Range(hide_range).Select
    Selection.EntireRow.Hidden = True
    Range(c2l(MSS_FIRST_SCORE_COL) & MSS_FIRST_PROJECT_ROW).Activate
    
    UnHideAllRowsColumns SHARED_SCORESHEET_TEMPLATE
    hide_range = c2l(SS_FIRST_RAW_COLUMN + num_criteria) & ":" & _
                 c2l(SS_FIRST_RAW_COLUMN + max_criteria - 1)
    hide_range = hide_range & "," & c2l(SS_FIRST_NORMAL_COLUMN + num_criteria) & ":" & _
                                    c2l(SS_FIRST_NORMAL_COLUMN + max_criteria - 1)
    hide_range = hide_range & "," & c2l(SS_FIRST_FINAL_COLUMN + num_criteria) & ":" & _
                                    c2l(SS_FIRST_FINAL_COLUMN + max_criteria - 1)
    Range(hide_range).Select
    Selection.EntireColumn.Hidden = True
    ' hide the unused rows at the bottom of the sheet
    hide_range = (SS_FIRST_DATA_ROW + num_projects * target_markers_per_proj) & ":" & _
                (SS_FIRST_DATA_ROW + max_projects * max_markers_per_proj)
    Range(hide_range).Select
    Selection.EntireRow.Hidden = True
    Range(c2l(SS_FIRST_RAW_COLUMN) & SS_FIRST_DATA_ROW).Activate
    
    UnHideAllRowsColumns MARKER_SCORING_TEMPLATE_SHEET
    hide_range = c2l(MST_FIRST_SCORING_COL + num_criteria) & ":" & _
                 c2l(MST_FIRST_SCORING_COL + max_criteria - 1)
    hide_range = hide_range & "," & c2l(MST_FIRST_NORMALIZED_SCORE_COLUMN + num_criteria) & ":" & _
                                    c2l(MST_FIRST_NORMALIZED_SCORE_COLUMN + max_criteria - 1)
    Range(hide_range).Select
    Selection.EntireColumn.Hidden = True
    Range(c2l(MST_FIRST_SCORING_COL) & MST_FIRST_SCORING_ROW).Activate

    UnHideAllRowsColumns MARKER_PROJECT_EXPERTISE_TEMPLATE
    hide_range = (MPET_FIRST_DATA_ROW + num_projects) & ":" & _
                 (MPET_FIRST_DATA_ROW + max_projects)
    Range(hide_range).Select
    Selection.EntireRow.Hidden = True
    Range(c2l(MPET_EXPERTISE_COLUMN) & MPET_FIRST_DATA_ROW).Activate
    
    UnHideAllRowsColumns MARKER_KEYWORD_EXPERTISE_TEMPLATE
    hide_range = (MKET_FIRST_DATA_ROW + num_keywords) & ":" & _
                 (MKET_FIRST_DATA_ROW + max_keywords - 1)
    Range(hide_range).Select
    Selection.EntireRow.Hidden = True
    Range(MKET_EXPERTISE_COLUMN & MKET_FIRST_DATA_ROW).Activate
       
    UnHideAllRowsColumns SCORES_AND_COMMENTS_INSTRUCTIONS_TEMPLATE
    hide_range = (SCI_PROJECT_COUNT_ROW + 2 + target_ass_per_marker) & ":" & _
                 (SCI_PROJECT_COUNT_ROW + 2 + max_ass_per_marker - 1)
    Range(hide_range).Select
    Selection.EntireRow.Hidden = True
    Range("A" & SCI_PROJECT_COUNT_ROW + 2).Activate
    
    UnHideAllRowsColumns SCORES_AND_COMMENTS_TEMPLATE_SHEET
    hide_range = (Range(SCT_FIRST_CRITERIA_SCORE).row + SCT_ROWS_PER_CRITERIA * num_criteria - 1) & ":" & _
                 (Range(SCT_FIRST_CRITERIA_SCORE).row + SCT_ROWS_PER_CRITERIA * max_criteria - 3)
    Range(hide_range).Select
    Selection.EntireRow.Hidden = True
    Range(SCT_COI_RESPONSE_CELL).Activate
    
    UnHideAllRowsColumns PROJECT_COMMENTS_SHEET
    hide_range = (Range(PC_GENERAL_COMMENTS_CELL).row + 1 + PC_NUM_ROWS_PER_COMMENT * num_criteria) & ":" & _
                 (Range(PC_GENERAL_COMMENTS_CELL).row + PC_NUM_ROWS_PER_COMMENT * max_criteria)
    Range(hide_range).Select
    Selection.EntireRow.Hidden = True
    Range(PC_GENERAL_COMMENTS_CELL).Activate
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    HideUnusedColumnsRowsF = True
End Function
    
Public Function UnhideUnusedColumnsRowsF() As Boolean

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    If (DefineGlobals = False) Then
        Exit Function
    End If
        
    UnHideAllRowsColumns PROJECT_KEYWORDS_SHEET
    UnHideAllRowsColumns MARKER_EXPERTISE_SHEET
    UnHideAllRowsColumns PROJECT_X_MARKER_SHEET
    UnHideAllRowsColumns EXPERTISE_CROSSWALK_SHEET
    UnHideAllRowsColumns MASTER_ASSIGNMENTS_SHEET
    UnHideAllRowsColumns MASTER_SCORESHEET
    UnHideAllRowsColumns SHARED_SCORESHEET_TEMPLATE
    UnHideAllRowsColumns MARKER_SCORING_TEMPLATE_SHEET
    UnHideAllRowsColumns MARKER_PROJECT_EXPERTISE_TEMPLATE
    UnHideAllRowsColumns MARKER_KEYWORD_EXPERTISE_TEMPLATE
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    UnhideUnusedColumnsRowsF = True
End Function

Public Function UnHideAllRowsColumns(sheet_name As String) As Boolean

    Dim start_address As String
    
    start_address = ActiveCell.Address
    Sheets(sheet_name).Select
    
    Cells.Select
    Selection.EntireColumn.Hidden = False
    Selection.EntireRow.Hidden = False
    
    Range(start_address).Select
    Range(start_address).Activate
    
    UnHideAllRowsColumns = True
End Function

Public Function Email2Text(str_in As String, max_char As Long) As String
    If max_char < 1 Then
        MsgBox "[Email2Text] maximum character length input as " & max_char & " ???", vbCritical
        Exit Function
    End If
    Dim i As Long, one_char As String
    For i = 1 To Len(str_in)
        one_char = Mid(str_in, i, 1)
        Select Case one_char
        Case "@"
            Email2Text = Email2Text & "_at_"
        Case Else
            If ((one_char >= "a") And (one_char <= "z")) Or _
                ((one_char >= "A") And (one_char <= "Z")) Or _
                ((one_char >= "0") And (one_char <= "9")) Then
                ' only keep the alphanumeric
                Email2Text = Email2Text & one_char
            End If
        End Select
    Next i
    If (Len(Email2Text) > max_char) Then
        Email2Text = Left(Email2Text, max_char)
    End If
End Function

Public Function FreeArrays() As Boolean

    FreeArrays = True
    Exit Function
    
    Erase mc_array, mc_as_loaded, pn_array, mn_array, coa_array, mentor_column
    Erase competition_COIs, ss_marker_col, ss_project_col, xlmh_per_marker, xlmh_per_project
    Erase assignments, n_assigned2project, n_assigned2marker, marker_orgs, marker_emails
    Erase assignment_failed_for_this_proj, messages, comments, general_comments

    FreeArrays = True
    
End Function
' callbacks from the macro buttons

Private Sub BuildMasterScoresheet_Click()
    Dim start_sheet As String, start_book As String
    start_sheet = ActiveSheet.Name
    start_book = ActiveWorkbook.Name
    InitMessages
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    If CompleteFinalScoresheets = False Then
        ReportMessages
    End If
    Workbooks(start_book).Activate
    Sheets(start_sheet).Activate
    ReportMessages
    FreeArrays
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = "All done."
End Sub

Private Sub ClearIntermediateAndOutputSheets_Click()
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    InitMessages
    Application.StatusBar = "Clearing sheets."
    ClearKeywordTablesPXMCrosswalkAssignmentsAndMss
    Application.StatusBar = "All done."
    ReportMessages
    FreeArrays
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub Expertise2Scoresheets_Click()
    Dim start_sheet As String
    start_sheet = ActiveSheet.Name
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    Expertise2MarkingSheets
    ThisWorkbook.Sheets(start_sheet).Activate
    FreeArrays
    Application.StatusBar = "All done."

End Sub

Private Sub Export_Click()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Dim start_sheet As String, start_book As String
    start_sheet = ActiveSheet.Name
    start_book = ActiveWorkbook.Name
    InitMessages
    If ExportCompetitionWorkbook = False Then
        ReportMessages
        Exit Sub
    End If
    ReportMessages
    Workbooks(start_book).Activate
    Sheets(start_sheet).Activate
    FreeArrays
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Private Sub FOM2Scoresheets_Click()
    InitMessages
    If KeywordTablesToScoresheets = False Then
    End If
    ReportMessages
    FreeArrays
End Sub

Private Sub GenerateExpertiseSheets_Click()
    Dim start_sheet As String, start_book As String
    start_sheet = ActiveSheet.Name
    start_book = ActiveWorkbook.Name
    InitMessages
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    
    MakeProjectExpertiseSheets
    
    ReportMessages
    Workbooks(start_book).Activate
    Sheets(start_sheet).Activate
    FreeArrays
    Application.StatusBar = "All done."
End Sub


Private Sub CreateKeywordExpertiseSheets_Click()
    Dim start_sheet As String, start_book As String
    start_sheet = ActiveSheet.Name
    start_book = ActiveWorkbook.Name
    InitMessages
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    
    MakeKeywordExpertiseSheets
    
    ReportMessages
    Workbooks(start_book).Activate
    Sheets(start_sheet).Activate
    FreeArrays
    Application.StatusBar = "All done."

End Sub

Private Sub HideUnusedColumnsRows_Click()
    Dim start_sheet As String, start_book As String
    start_sheet = ActiveSheet.Name
    start_book = ActiveWorkbook.Name
    InitMessages
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    If HideUnusedColumnsRowsF = False Then
        ReportMessages
        Exit Sub
    End If
    Workbooks(start_book).Activate
    Sheets(start_sheet).Activate
    ReportMessages
    FreeArrays
    Application.StatusBar = "All done."
End Sub

Private Sub LoadExpertiseByKeyword_Click()
    Dim start_sheet As String
    start_sheet = ActiveSheet.Name
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    InitMessages
    If LoadMarkerKeywordExpertiseIntoPXM = False Then
        ReportMessages
        Exit Sub
    End If
    If CreatePXMFromProjectRelevanceAndMarkerExpertise = False Then
        ReportMessages
        Exit Sub
    End If
    ReportMessages
    Sheets(start_sheet).Activate
    FreeArrays
    Application.StatusBar = "All done."

End Sub

Private Sub LoadPerProjectExpertise_Click()
    Dim start_sheet As String
    start_sheet = ActiveSheet.Name
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    InitMessages
    If LoadMarkerProjectExpertiseIntoPXM = False Then
        ReportMessages
        Exit Sub
    End If
    ReportMessages
    Sheets(start_sheet).Activate
    FreeArrays
    Application.StatusBar = "All done."
End Sub

Private Sub MakeAssignments_Click()
    Dim start_sheet As String
    start_sheet = ActiveSheet.Name
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    InitMessages
    If AssignMarkersBasedOnConfidence = False Then
        ReportMessages
        Exit Sub
    End If
    If PopulateSharedScoresheet = False Then
        ReportMessages
        Exit Sub
    End If
    ReportMessages
    Sheets(start_sheet).Activate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    FreeArrays
    Application.StatusBar = "All done."
End Sub

Private Sub CreateMarkerScoresheets1_Click()
    Dim start_sheet As String
    start_sheet = ActiveSheet.Name
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    InitMessages
    If CreateAllMarkingSheets = False Then
        ReportMessages
        Exit Sub
    End If
    ReportMessages
    Sheets(start_sheet).Activate
    FreeArrays
    Application.StatusBar = "All done."

End Sub

Private Sub CalculateFOM_Click()
    Dim start_sheet As String
    start_sheet = ActiveSheet.Name
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    InitMessages
    If (CreateFigureOfMeritTable = False) Then
        ReportMessages
        Exit Sub
    Else
        AddMessage "Finished combining Project and Expertise keyword ratings for " & num_projects & " projects, " & _
                    num_markers & " markers."
    End If
    ReportMessages
    Sheets(start_sheet).Activate
    FreeArrays
    Application.StatusBar = "All done."
End Sub

Private Sub ShowAllRowsAndColumns_Click()
    Dim start_sheet As String
    start_sheet = ActiveSheet.Name
    Application.StatusBar = "Processing - wait for pop-up dialog at end of processing."
    InitMessages
    If UnhideUnusedColumnsRowsF = False Then
        ReportMessages
        Exit Sub
    End If
    ReportMessages
    Sheets(start_sheet).Activate
    FreeArrays
    Application.StatusBar = "All done."
    
End Sub