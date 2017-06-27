function Add-HTMLTableColor {
    <# 
    .SYNOPSIS 
    Colorize cells or rows in an HTML table, or add other inline CSS
 
    .DESCRIPTION 
    Colorize cells or rows in an HTML table, or add other inline CSS
 
    .PARAMETER  HTML 
    HTML string to work with

    .PARAMETER  Column 
    If specified, the column you want to modify.  This is case sensitive

    .PARAMETER  Argument 
    If Column is specified, this argument can be used to compare with current cell.

    .PARAMETER ScriptBlock
    If Column is specified, used to evaluate whether to colorize a cell.  If the scriptblock returns $true the cell will be colorized.
   
    $args[0] is the existing cell value in the table
    $args[1] is your Argument parameter

    Examples:
        {[string]$args[0] -eq [string]$args[1]} #existing cell value equals Argument.  This is the default
        {[double]$args[0] -gt [double]$args[1]} #existing cell value is greater than Argument.

    Use strong typesetting if possible.
 
    .PARAMETER  Attr 
    If Column is specified, the attribute to change should ColumnValue be found in the Column specified or if the ScriptBlock is true.  Default:  Style
 
    .PARAMETER  AttrValue 
    If Column is specified, the attribute value to set when the ColumnValue is found in the Column specified or if the ScriptBlock is true.
    
    Example: "background-color:#FFCC99;" 
 
    .PARAMETER WholeRow
    If specified, and Column is specified, set the Attr and AttrValue for the entire row, not just a cell.

    .EXAMPLE
    #This example requires and demonstrates using the New-HTMLHead, New-HTMLTable, Add-HTMLTableColor, ConvertTo-PropertyValue and Close-HTML functions.
    
    #get processes to work with
        $processes = Get-Process
    
    #Build HTML header
        $HTML = New-HTMLHead -title "Process details"

    #Add CPU time section with top 10 PrivateMemorySize processes.  This example does not highlight any particular cells
        $HTML += "<h3>Process Private Memory Size</h3>"
        $HTML += New-HTMLTable -inputObject $($processes | sort PrivateMemorySize -Descending | select name, PrivateMemorySize -first 10)

    #Add Handles section with top 10 Handle usage.
    $handleHTML = New-HTMLTable -inputObject $($processes | sort handles -descending | select Name, Handles -first 10)

        #Add highlighted colors for Handle count
            
            #build hash table with parameters for Add-HTMLTableColor.  Argument and AttrValue will be modified each time we run this.
            $params = @{
                Column = "Handles" #I'm looking for cells in the Handles column
                ScriptBlock = {[double]$args[0] -gt [double]$args[1]} #I want to highlight if the cell (args 0) is greater than the argument parameter (arg 1)
                Attr = "Style" #This is the default, don't need to actually specify it here
            }

            #Add yellow, orange and red shading
            $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 1500 -attrValue "background-color:#FFFF99;" @params
            $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 2000 -attrValue "background-color:#FFCC66;" @params
            $handleHTML = Add-HTMLTableColor -HTML $handleHTML -Argument 3000 -attrValue "background-color:#FFCC99;" @params
      
        #Add title and table
        $HTML += "<h3>Process Handles</h3>"
        $HTML += $handleHTML

    #Add process list containing first 10 processes listed by get-process.  This example does not highlight any particular cells
        $HTML += New-HTMLTable -inputObject $($processes | select name -first 10 ) -listTableHead "Random Process Names"

    #Add property value table showing details for PowerShell ISE
        $HTML += "<h3>PowerShell Process Details PropertyValue table</h3>"
        $processDetails = Get-process powershell_ise | select name, id, cpu, handles, workingset, PrivateMemorySize, Path -first 1
        $HTML += New-HTMLTable -inputObject $(ConvertTo-PropertyValue -inputObject $processDetails)

    #Add same PowerShell ISE details but not in property value form.  Close the HTML
        $HTML += "<h3>PowerShell Process Details object</h3>"
        $HTML += New-HTMLTable -inputObject $processDetails | Close-HTML

    #write the HTML to a file and open it up for viewing
        set-content C:\test.htm $HTML
        & 'C:\Program Files\Internet Explorer\iexplore.exe' C:\test.htm

    .EXAMPLE
    # Table with the 20 most recent events, highlighting error and warning rows

        #gather 20 events from the system log and pick out a few properties
        $events = Get-EventLog -LogName System -Newest 20 | select TimeGenerated, Index, EntryType, UserName, Message

    #Create the HTML table without alternating rows, colorize Warning and Error messages, highlighting the whole row.
        $eventTable = $events | New-HTMLTable -setAlternating $false |
            Add-HTMLTableColor -Argument "Warning" -Column "EntryType" -AttrValue "background-color:#FFCC66;" -WholeRow |
            Add-HTMLTableColor -Argument "Error" -Column "EntryType" -AttrValue "background-color:#FFCC99;" -WholeRow

    #Build the HTML head, add an h3 header, add the event table, and close out the HTML
        $HTML = New-HTMLHead
        $HTML += "<h3>Last 20 System Events</h3>"
        $HTML += $eventTable | Close-HTML

    #test it out
        set-content C:\test.htm $HTML
        & 'C:\Program Files\Internet Explorer\iexplore.exe' C:\test.htm

    .NOTES 
    Props to Zachary Loeber and Jaykul for the idea and help:
    http://gallery.technet.microsoft.com/scriptcenter/Colorize-HTML-Table-Cells-2ea63acd
    http://stackoverflow.com/questions/4559233/technique-for-selectively-formatting-data-in-a-powershell-pipeline-and-output-as

    I believe that .Net 3.5 is a requirement for using the Linq libraries
    
    .FUNCTIONALITY
    General Command
    #> 
    [CmdletBinding()] 
    param ( 
        [Parameter( Mandatory = $true,  
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $false)]  
        [string]$HTML,
        
        [Parameter( Mandatory = $false, 
            ValueFromPipeline = $false)]
        [String]$Column = "Name",
        
        [Parameter( Mandatory = $false,
            ValueFromPipeline = $false)]
        $Argument = 0,
        
        [Parameter( ValueFromPipeline = $false)]
        [ScriptBlock]$ScriptBlock = {[string]$args[0] -eq [string]$args[1]},
        
        [Parameter( ValueFromPipeline = $false)]
        [String]$Attr = "style",
        
        [Parameter( Mandatory = $true, 
            ValueFromPipeline = $false)] 
        [String]$AttrValue,
        
        [Parameter( Mandatory = $false, 
            ValueFromPipeline = $false)] 
        [switch]$WholeRow = $false

    )
    
    #requires -version 2.0
    add-type -AssemblyName System.xml.linq | out-null

    # Convert our data to x(ht)ml  
    $xml = [System.Xml.Linq.XDocument]::Parse($HTML)   
        
    #Get column index.  try th with no namespace first, then default namespace provided by convertto-html
    try { 
        $columnIndex = (($xml.Descendants("th") | Where-Object { $_.Value -eq $Column }).NodesBeforeSelf() | Measure-Object).Count 
    }
    catch { 
        Try {
            $columnIndex = (($xml.Descendants("{http://www.w3.org/1999/xhtml}th") | Where-Object { $_.Value -eq $Column }).NodesBeforeSelf() | Measure-Object).Count
        }
        Catch {
            Throw "Error:  Namespace incorrect."
        }
    }

    #if we got the column index...
    if ($columnIndex -as [double] -ge 0) {
            
        #take action on td descendents matching that index
        switch ($xml.Descendants("td") | Where-Object { ($_.NodesBeforeSelf() | Measure-Object).Count -eq $columnIndex }) {
            #run the script block.  If it is true, set attributes
            {$(Invoke-Command $ScriptBlock -ArgumentList @($_.Value, $Argument))} { 
                    
                #mark the whole row or just a cell depending on param
                if ($WholeRow) { 
                    $_.Parent.SetAttributeValue($Attr, $AttrValue) 
                } 
                else { 
                    $_.SetAttributeValue($Attr, $AttrValue) 
                }
            }
        }
    }
        
    #return the XML
    $xml.Document.ToString() 
}