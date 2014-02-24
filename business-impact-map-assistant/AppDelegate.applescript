--
--  AppDelegate.applescript
--  Business Impact Map Assistant
--
--  Created by Tommy on 2014-02-15.
--  Copyright (c) 2014 Helt Enkelt AB. All rights reserved.
--
--
--
--
--

script AppDelegate
	property parent : class "NSObject"
	
	property red : {65535, 10023, 4626}
	property white : {65535, 65535, 65535}
	property black : {0, 0, 0}
	property grey : {40000, 40000, 40000} ---{60000, 60000, 60000}
    property darkgrey : {50000, 50000, 50000}
	
	
	on applicationWillFinishLaunching_(aNotification)
		-- Insert code here to initialize your application before any files are opened 
	end applicationWillFinishLaunching:
	
	on applicationShouldTerminate_(sender)
		-- Insert code here to do any housekeeping before your application quits 
		return current application's NSTerminateNow
	end applicationShouldTerminate:
	
	on generateFacetTables_(sender)
		my prepareForGenerateImpactTableBy()
        my generateImpactTableBy("Syfte")
		my generateImpactTableBy("Målgrupp")
		my generateImpactTableBy("Användningsmål")  
		my generateImpactTableBy("Åtgärd")          
	end generateAspectTables_
	
    on prepareForGenerateImpactTableBy()
		tell application "Microsoft Excel"
            set sourceSheet to sheet "Master" of active workbook
            tell active workbook
                -- Remove any existing generated sheets
                ---my deleteSheet2({"(cache)", "Syfte", "Målgrupp", "Användningsmål", "Åtgärd"})
                
                my deleteSheet("(cache)")
                my deleteSheet("Syfte")            
                my deleteSheet("Målgrupp")
                my deleteSheet("Användningsmål")
                my deleteSheet("Åtgärd")  
                
                -- Create a cache sheet
				set lastSheet to sheet (count of sheets)
				copy worksheet sourceSheet after lastSheet
				tell sheet (count of sheet) of active workbook					
                    set name to "(cache)"					
				end tell
                set currentSheet to sheet (count of sheets)
                
                -- Format cache sheet, fill blank cells, etc.
                
                --      remove empty columns
                ----my deleteEmptyColumns(currentSheet)
                
                --      remove empty rows
                my deleteEmptyRows(currentSheet)  
                ---display dialog "Kolla ute till höger"
                
                --      remove lines
                my clearLines(currentSheet)
                    
                --      fill empty cells to make them sortable
                my fillInTheBlanks(currentSheet) 
            end
        end
    end
    
	on generateImpactTableBy(facet)
		tell application "Microsoft Excel"
			set sourceSheet to sheet "(cache)" of active workbook
			tell active workbook
                -- Remove earlier generation, if any
				---my deleteSheet(facet)
				
                -- Copy cache
				set lastSheet to sheet (count of sheets)
				copy worksheet sourceSheet after lastSheet
				tell sheet (count of sheet) of active workbook					
                    set name to facet					
				end tell
                set currentSheet to sheet (count of sheets)
                
                tell currentSheet
                    -- Reset text color
                    ---set theUsedRange to used range
                    ---set color of font object of theUsedRange to black
                    
                    --

                    my sortBy(currentSheet, facet)
				end
                

                -- Undo selection
                select cell "A1"
			end tell
		end tell
	end generateImpactTableBy
	
    -- Sorts on the primary column
	on sortBy(currentSheet, primary)
        tell application "Microsoft Excel"
            tell currentSheet  
                set noOfColumns to 1       -- default. Numbers of columns under one facet, used when columns are moved
                if primary is "Syfte" then 
                    set keyRange to range("A:A")
                    set colIndex to 1
                else 
                    if primary is "Målgrupp" then 
                        set keyRange to range("B:B")
                        set colIndex to 2
                    else
                        if primary is "Användningsmål" then 
                            set keyRange to range("C:C")
                            set colIndex to 3
                        else
                            if primary is "Åtgärd" then 
                                set keyRange to range("D:D")
                                set colIndex to 4
                                set noOfColumns to 5

                            else
                                -- Not any major category
                                error primary & " is a unknown column. Can not sort on it."
                            end
                        end
                    end
                end
                -- TODO fortsätt här
                
                -- TODO remove lines
                ---my clearLines(currentSheet)
                
                -- Fill empty cells to make them sortable
                ---my fillInTheBlanks(currentSheet)                  
                
                -- Sort
                select keyRange
                sort used range of currentSheet key1 (keyRange) header header yes
                ---sort used range of currentSheet key1 (range "A:A" of worksheet currentSheet) header header yes
                -- TODO Add sensable secondary keys.
                
                -- Move sorted column to leftmost
                my moveColumnsToFarLeft(currentSheet, colIndex, noOfColumns)
                
                -- Todo Empty duplicate cells
                my makeDuplicatesBlankAddLines(currentSheet)                    
                
                -- Remove columns without content
                my deleteEmptyColumns(currentSheet)                
                
            end
            tell active window to set display gridlines to false                
        end
    end
    
    -- Move one or more column to the leftmost position
    -- colIndex = index of the column (or the leftmost of the columns)
    -- noOfColumns = Number of columns to move
    on moveColumnsToFarLeft(currentSheet, colIndex, noOfColumns)
        tell application "Microsoft Excel"
            tell currentSheet
                set leftmostLetter to my indexToChar(colIndex)
                set rightmostLetter to my indexToChar(colIndex + noOfColumns - 1)
                set countLetter to my indexToChar(noOfColumns)
                set adjustedLeftmostLetter to my indexToChar(colIndex + noOfColumns)
                set adjustedRightmostLetter to my indexToChar(colIndex + (noOfColumns * 2) - 1)
                
                insert into range entire column of column ("A:" & countLetter) -- A blank area wide enough to host the column(s) that are going to move there
                cut range (range (adjustedLeftmostLetter & ":" & adjustedRightmostLetter)) destination of cut range ("A1")
                delete range entire column of range (adjustedLeftmostLetter & ":" & adjustedRightmostLetter)
            end
        end
    end
    
    -- Helper that translates a column index into a letter (1 = A, 2 = B, etc) 
    on indexToChar(index)
        return character index of "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    end indexToChar
    
    -- Before sorting, the empty cells (signifying that they have the value of the cell above) must be made explicit
    on fillInTheBlanks(currentSheet)
        tell application "Microsoft Excel"
            tell currentSheet
                repeat with c from 1 to count of columns of used range
                    set previousCell to false -- Normaly value of cell above is checked, but this is not applicable on top of columen.
                    repeat with r from 2 to count of rows of used range -- (Row 1 is for headings)

                        set currentCell to my aCell(currentSheet, r, c)
                        select currentCell
                        
                        -- BUG? DET HÄR GÅR VÄL FRÅN VÄRDET PÅ NEDERSTA TILL ÖVERSTA PÅ NÄSTA SPALT NU?
                        if value of currentCell as string is "" then
                            -- Empty cell, may need to be filled
                            if r is 1 then
                               -- Top of column. Nothing to fill with. Leave blank.  
                            end
                            if r > 2 then
                                -- Need a cell (other then header) above to collect value from
                                if previousCell is not false then                                    
                                    --- and value of previousCell as string is not "" and value of previousCell as string is not "---" then
                                    -- Need something reasonable to copy
                                    
                                    -- Härnäst: Om cellen -1 -1 inte är samma som cellen 0 -1, så kopiera inte
                                    if c > 1 then                                        
                                        set preParentCell to my aCell(currentSheet, r - 1, c - 1)
                                        set parentCell to my aCell(currentSheet, r, c - 1)
                                        if value of preParentCell as string is value of parentCell as string then
                                        -- previous and current cell are both of the same parent
                                    
                                    -- Och (fast här behöver jag tänka mer) Om cellvärdet INTE är detsamma som i cellen ovanför, ska den och alla celler till höger om den, ha en linje ut till högerkanten. Tjockare/svartare linje ju längre ut till vänster man börjar (4 nivåer behövs)
                                    -- Problem: Blir randigt längst ut till höger. Inga linjer startar till höger om första hur. Inga linjer dras ut om det bara är tomt (eller ---) i alla celler till höger om dem
                                    
                                        --display dialog value of previousCell as string
                                            set value of currentCell to value of previousCell as string
                                            set color of font object of currentCell to grey
                                        end
                                    else
                                        -- (No parent cell, so no need to check for it)
                                        set value of currentCell to value of previousCell as string
                                        set color of font object of currentCell to grey
                                    end
                                end if
                            end if
                        end if
                        
                        (*
                        -- Grey out cells that has the same value as the cell above
                        if value of currentCell as string is not "" and previousCell is not false then
                            if value of currentCell as string is value of previousCell as string then
                                set color of font object of currentCell to grey
                                else
                                -- Current cell is not the same as the cell above it
                                set weight of (get border of currentCell which border edge top) to border weight thin
                            end if
                        end if
                         *)
                        
                        set previousCell to currentCell
                        
                        
                    end repeat
                end repeat
            end
        end
	end
    
    -- To make readability better, cells whose value is a copy of the cell above should be empty. 
    on makeDuplicatesBlankAddLines(currentSheet)        
        tell application "Microsoft Excel"
                        tell active window to set display gridlines to false 
            tell currentSheet
                
                -- If cell is empty, copy value from above
                ---set previousCell to false

                
                repeat with c from 8 to 1 by -1 
                    set currentTopic to false
                    repeat with r from 2 to count of rows of used range -- (Row 1 is for headings)
                        
                        set currentCell to my aCell(currentSheet, r, c)
                        select currentCell
                        set color of font object of currentCell to black                        
                        
                        if r is 2 then
                            set currentTopic to value of currentCell as string 
                        end
                        
                        if r > 2 then
                            if value of currentCell as string is not currentTopic
                                -- The start of a new topic
                                set currentTopic to value of currentCell as string
                                                                
                                --TODO START LINJE
                                ---set weight of (get border of currentCell which border edge top) to border weight thin
                                
                                -- Draw lines to help the eye
                                if c <= 8 --4  --MÅSTE VARA C     -- (Don't want to make the lines to detailed  TODO THIS WILL BE STRANGE ON WHEN WHAT IS TO THE LEFT)
                                    set rangeThatNeedsBorder to range my getAllCellsRightwards(currentSheet, r, c)
                                    set theBorder to (get border of rangeThatNeedsBorder which border edge top)
                                    set weight of theBorder to border weight thin
                                    set color of theBorder to darkgrey
                                end
                            else
                                -- Same topic as cell above
                                ---- TEST TILLFÄLLIGT BORTKOPPLAD set value of currentCell to ""
                                --- TEST ISTÄLLET GÖRS TEXTEN GRÅ
                                ---set color of font object of currentCell to grey
                                -- TEST CITATTECKEN
                                if currentTopic is not ""
                                    set value of currentCell to "\""
                                    set vertical alignment of currentCell to valign center
                                    set horizontal alignment of currentCell to horizontal align center
                                end
                           end                        
                        end
                    
                    ----    
                      
                                    -- Need something reasonable to copy
                                    
                                    -- Härnäst: Om cellen -1 -1 inte är samma som cellen 0 -1, så kopiera inte
                                    
                                    -- Och (fast här behöver jag tänka mer) Om cellvärdet INTE är detsamma som i cellen ovanför, ska den och alla celler till höger om den, ha en linje ut till högerkanten. Tjockare/svartare linje ju längre ut till vänster man börjar (4 nivåer behövs)
                                    -- Problem: Blir randigt längst ut till höger. Inga linjer startar till höger om första hur. Inga linjer dras ut om det bara är tomt (eller ---) i alla celler till höger om dem                          
                    end repeat
                end repeat
            end
        end
    end
    
	-- IBActions (button clicks)
	on generateImpactTables_(sender)
		tell application "Microsoft Excel"
			set sourceSheet to sheet "Master" of active workbook
			tell active workbook
				set lastSheet to sheet (count of sheets)
			end tell
			set aw to active workbook
			copy worksheet sourceSheet after lastSheet
			tell sheet (count of sheet) of active workbook
				set name to "GENERERAD"
				--copy worksheet sheet "Översikt" of active workbook after sheet 1
				
				
				
				--tell application "Microsoft Excel"
				--		set theUsedRange to used range of sheet "Översikt" of active workbook
				--		set color of font object of theUsedRange to black
				--		get value of range "B3" of sheet "Översikt" of active workbook
				--	end tell
				
				-- Reset text color
				set theUsedRange to used range
				set color of font object of theUsedRange to black
				--get value of range "B3" of sheet "Översikt" of active workbook
				--display dialog value of range "B2" of sheet "Översikt" of active workbook as string
				
				
				-- Move to new sheet
				tell application "Microsoft Excel"
					
				end tell
				
				
				
				
				-- TODO Ta bort helt tomma rader
				
				
				-- If cell is empty, copy value from above
				set previousCell to false
				
				repeat with c from 1 to count of columns of used range
					repeat with r from 2 to count of rows of used range -- (Row 1 is for headings)
						
						set currentCell to my aCell(sourceSheet, r, c)
						
						
						if value of currentCell as string is "" then
							-- Empty cell
							if r > 2 then
								-- Need a cell (other then header) above to collect value from
								if previousCell is not false and value of previousCell as string is not "" and value of previousCell as string is not "---" then
									-- Need something reasonable to copy
									
									-- Härnäst: Om cellen -1 -1 inte är samma som cellen 0 -1, så kopiera inte
									
									-- Och (fast här behöver jag tänka mer) Om cellvärdet INTE är detsamma som i cellen ovanför, ska den och alla celler till höger om den, ha en linje ut till högerkanten. Tjockare/svartare linje ju längre ut till vänster man börjar (4 nivåer behövs)
									-- Problem: Blir randigt längst ut till höger. Inga linjer startar till höger om första hur. Inga linjer dras ut om det bara är tomt (eller ---) i alla celler till höger om dem
									
									--display dialog value of previousCell as string
									set value of currentCell to value of previousCell as string
								end if
							end if
						end if
						
						-- Grey out cells that has the same value as the cell above
						if value of currentCell as string is not "" and previousCell is not false then
							if value of currentCell as string is value of previousCell as string then
								set color of font object of currentCell to grey
                                else
								-- Current cell is not the same as the cell above it
								set weight of (get border of currentCell which border edge top) to border weight thin
							end if
						end if
						
						set previousCell to currentCell
						
						
					end repeat
				end repeat
				
				
			end tell
		end tell
		
		-- TODO (Troligen i ett annat script, som dock behöver köra det här efteråt.
		
		-- Sortera på olika saker, och gruppera sedan. Verkar som att enda sättet att få en "*** rubrik ***" på sin sortering är att låta scriptet sätta dit den, ovanför gruppen (verkar knäppt, måste kolla mer). 
		
		-- Jag funderade ett tag på att flytta spaltordningen så det sorterade hamnade längst ut till vänster, men det blir nog bara knäppt. Däremot kanske de svarta rubrikerna i den kolumnen ska markeras ännu mer, genom att bli feta, eller genom att bli röda etc.
		
		-- Måste fundera över om logiken med linjerna funkar även efter en sortering...
		
		-- Alternativt producerar man olika sheets för olika saker. Men det blir vääääldigt många. 
		
		-- Note bene, när man väl börjat sortera kommer man inte tillbaka till ursprungsläget		
	end generateImpactTables_
	
	-- Returns a cell, given numeric coordinates (row first!)
	-- For example, returns the cell on "B3" from input 3, 2
	on aCell(theSheet, row_no, col_no)
		set cLetter to character col_no of "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		set addr to cLetter & row_no
		tell application "Microsoft Excel"
			tell theSheet --of active workbook
				set theCell to range (addr)
			end tell
		end tell
		return theCell
	end aCell
    
    -- Deletes sheet named theName, if it exist. 
	on deleteSheet(theName)
        tell application "Microsoft Excel"
            set wsheets to worksheets whose name is theName
            if wsheets is missing value then
                -- No worksheet with that name. Do nothing
                else
				tell active workbook
					activate
					delete first item of wsheets
				end tell
            end if
        end tell
	end deleteSheet
    on deleteSheet2(theNames)
        tell application "Microsoft Excel"
            set sheetsToDelete to {}
            repeat with aName in theNames
                set maybeWsheet to worksheets whose name is aName
                if maybeWsheet is missing value then
                    -- No worksheet with that name. Do nothing
                else
                    set sheetsToDelete  to sheetsToDelete & (first item of maybeWsheet) 
                end                
            end
            tell active workbook
                activate
                delete sheetsToDelete     
            end
        end tell
    end deleteSheet2


    -- Remove all lines from the worksheet
    on clearLines(currentSheet)
        tell application "Microsoft Excel"
            tell currentSheet  
                repeat with r from 2 to count of rows of used range                    
                    set currentRow to range ("A" & r & ":H" & r)       -- Only seams to work reliably on one row at a time
                    select currentRow
                    set line style of (get border of currentRow which border edge top) to None
                end
            end
        end
    end 

    -- Remove empty rows
    on deleteEmptyRows(currentSheet)
        tell application "Microsoft Excel"
            tell currentSheet  
                repeat with r from (count of rows of used range) to 2 by -1
                    set currentRow to range ("A" & r & ":H" & r)
                    select currentRow
                    if (value of cells of currentRow) as string is "" then
                        -- Empty row. Delete it. 
                        delete range (currentRow) shift shift up
                    end if
                end repeat
                
            end
        end
    end

    -- Remove empty columns
    on deleteEmptyColumns(currentSheet)
        tell application "Microsoft Excel"
            tell currentSheet  
                repeat with c from (count of columns of used range) to 1 by -1
                    set noRows to (count of rows of used range)
                    set columnLetter to character c of "ABCDEFGHIJKLMNOPQRSTUVWXYZ" 
                    set currentColumn to range (columnLetter & "1:" & columnLetter & noRows)
                    select currentColumn
                    if (value of cells of currentColumn) as string is "" then
                        -- Empty row. Delete it. 
                        delete range (entire column of currentColumn)
                    end if
                end repeat
            end
        end
    end

    -- Returns a range with current cell and all cells to the right of it
    on getAllCellsRightwards(currentSheet, contextR, contextC)
        tell application "Microsoft Excel"
            tell currentSheet
                set leftmostLetter to character contextC of "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                set leftAddr to leftmostLetter & contextR
                set rightmostLetter to character (count of columns of used range) of "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                set rightAddr to rightmostLetter & contextR
                
                tell currentSheet
                    set theRange to leftAddr & ":" & rightAddr
                end tell
                return theRange
            end tell
        end tell
    end getAllCellsRightwards

    
    -- Från http://macscripter.net/viewtopic.php?id=33221
    -- OBS, kolla kommentarerna. Verkar finnas bättre sätt där!
    -- Exempel:
    -- columnLetterBase26(26) --> "z"
    -- columnLetterBase26(27) --> "aa"
    -- columnLetterBase26(28) --> "ab"
    -- columnLetterBase26(289) --> "kc"
    (*
    on columnLetterBase26(aDecimal)
        try
            set aDecimal to aDecimal as integer
            on error
            error "The parameter given to convert decimals to base 26 wasn't a number." from aDecimal to integer
        end try
        
        if aDecimal < 1 then error "The parameter given is smaller than 1" from aDecimal to integer
        
        local Base26
        set Base26 to "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        
        -- search for largest exponent
        set exponent to 0
        repeat while (26 * (26 ^ exponent)) < aDecimal
            set exponent to exponent + 1
        end repeat
        
        -- fill string
        set repr26 to {}
        repeat while exponent ≥ 0
            
            -- find biggest suitable
            set baseNumber to 26
            repeat while (baseNumber * (26 ^ exponent)) > aDecimal
                set baseNumber to baseNumber - 1
            end repeat
            
            -- add letter
            set end of repr26 to item baseNumber of Base26
            
            -- addapt values
            set aDecimal to aDecimal - (baseNumber * (26 ^ exponent))
            set exponent to exponent - 1
            
        end repeat
        exponent
        
        
        return (repr26 as string)
    end columnLetterBase26
    *)

    (*
    -- See http://www.mactech.com/articles/mactech/Vol.23/23.02/2302AppleScript/index.html
    on isColumnEmpty
    end
    on isRowEmpty
    end
    *)    

end script






-- Make the sheet easier to read, by drawing lines on apropriate places
on drawLines()
	tell application "Microsoft Excel"
		tell sourceSheet
			set previousCell to false
			repeat with c from 4 to 1 by -1 -- The fouth column is the Hur/How. Subdividing with lines to the right
				-- of that will just be visually overhelming. Going from the right to the left makes it possible
				-- to make the line thicker the more to the left it has it's origin.
				repeat with r from 2 to (count of rows of used) -- range) -- (Row 1 is for headings)
					set currentCell to my aCell(r, c)
					if r > 2 then
						if currentCell as string is not previousCell as string and currentCell as string is not "" then
							-- A change of value => Draw the line
							set rangeThatNeedsBorder to my getAllCellsRightwards(sourceSheet, r, c)
							set weight of (get border of rangeThatNeedsBorder which border edge top) to border weight thin -- FUNKAR INTE, TROTS ATT SAMMA RAD ACCEPTERAS INNE I GENERATE IMPACT TABLES.  
						end if
					end if
					set value of previousCell to value of currentCell as string -- (Note, this will be somewhat strange
					-- when r = 1 or 2, but since those lines are not used, it does not matter)
				end repeat
			end repeat
		end tell
	end tell
end drawLines



-- Check if there already is a line
-- Not implemented

