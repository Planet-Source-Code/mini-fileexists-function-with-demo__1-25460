Attribute VB_Name = "FileExist"
        '###########################################################
        
        '###########################################################
        '#              This is a function that will               #
        '#      check if the file is present on the gived Path     #
        '###########################################################
        
            'This is how to call the function on the Form_Load() :
            
            'Private Sub Form_Load()
            'Call FileExist(App.Path & "\Data.dat")
                        'Or
            'Call FileExist(C:\SCANDISK.LOG")
            'End Sub
        
        '###########################################################
        '#          I wish that it will help sumebody              #
        '#     I think this is one of the easiest way to do it     #
        '###########################################################
        
        '###########################################################
        
Option Explicit

Public Function FileExists(FileName As String) As Boolean

    Dim Variable As String       'Variable for this module.
    
    On Error GoTo NotThere       'Simulate the occurrence of an error.
       Variable = Dir$(FileName) 'send back a string value.
    If Variable = "" Then        'If the Variable = Nothing then FileExist will = 0.
       FileExists = False        'False = 0
       MsgBox "the file is missing", "***** Erreur 69 *****"
       End                       'Close the execution immediately.
    Else
        FileExists = True        'True = 1
    End If                       'If Variable = "" Then
    
NotThere:                        'The error reference.
    If Err = 53 Then Resume Next 'If the Simulate Error occure then will resume next.

End Function

