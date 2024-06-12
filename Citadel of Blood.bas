    global gNoSegBmps, gNoMiscBmps, gFeatures$, gNoFeatures,_
           gViewRows, gViewCols, gViewRowStart, gViewColStart,_
           gGridRowsInc, gGridColsInc, gGridNoRows, gGridNoCols,_
           gPartyGridRow, gPartyGridCol, gCurrentLevel,_
           gHellgateFound, gHellgateBmpPrefix$, gHellgateLevel, gHellgateSegsAway, gUnvisitedBmp$,_
           gCorridor$, gDoor$, gWall$, gRoom$, gUnknown$, gNormal$, gPartyLoc$,_
           gGameChanged, gOpenfile$,_
           gFirstMirrorFound, gSecondMirrorFound

    call initGlobals
    call initArrays
    call loadBitmaps
    call createButtons  ' Separate sub so can be put at bottom of this file out of the way (it's big)
    call openMainWnd

    ' ------------------------------------------------------------------------------------------------------------
    '
    ' Start-up close-down routines
    '
    ' ------------------------------------------------------------------------------------------------------------

    ' Called first time program loads
    sub initGlobals
        gFeatures$ = "Altar Artwork Fountain Furniture Mirror Staircase Statue Trap Empty Hellgate Gateway"
        gNoFeatures = 11
        gNoSegBmps = 0      ' Number of segment (corridor and room) bitmap names stored in segBmps$() array
                            ' the 64 bit party location versions
        gNoMiscBmps = 0     ' Remaining segment (corridor and room) bitmap names stored in miscBitmaps$() array
                            ' the 64 bit non party location versions, all 48 and 32 bit versions,
                            ' unvisited location, and Gateway bitmaps
        gViewRows = 15      ' Width of grid viewport visible to user
        gViewCols = 25      ' Height of grid viewport visible to user
        gGridRowsInc = 7    ' How much to expand no of grid rows when party reaches viewport edge
        gGridColsInc = 12   ' How much to expand no of grid cols when party reaches viewport edge
        gRoom$ = "r"        ' Segment tile type
        gCorridor$ = "c"    ' Can be segment tile type or segment exit type
        gDoor$ = "d"        ' Segment exit type
        gWall$ = "w"        ' Segment (non) exit type
        gUnknown$ = "u"     ' Used with Hellgate and Unvisited segments for tile type and exit type
        gNormal$ = "n"      ' Bitmap name 10th char denotes a bitmap showing party NOT at that location
        gPartyLoc$ = "p"    ' Bitmap name 10th char denotes a bitmap showing party at that location
        gUnvisitedBmp$ = "uuuuu"; getFeatureNo$("Empty"); gNormal$; "64"

        call initGlobalsOnNew
    end sub

    ' Called each time a new game started
    sub initGlobalsOnNew
        gOpenfile$ = ""     ' Currently open file for save menu option
        gGameChanged = 0    ' To check if user wants to save on program close
        gViewRowStart = 1   ' First row of grid in viewport
        gViewColStart = 1   ' First column of grid in viewport
        gCurrentLevel = 1   ' Which of 3 levels party is located in
        gHellgateFound = 0  ' Only one Hellgate in game
        gFirstMirrorFound = 0
        gSecondMirrorFound = 0
        gHellgateSegsAway = -1
        gHellgateBmpPrefix$ = ""    ' Once hellgate located unknown exits randomly assigned
        gGridNoRows = gViewRows     ' Grid of segments starts same size as viewport but grows
        gGridNoCols = gViewCols     ' as dungeon extends past initial viewport boundaries
        gPartyGridRow = int(gViewRows/2)
        gPartyGridCol = int(gViewCols/2)
    end sub

    ' This program use array indexes from 1 so dimensions are 1 larger than required number of elements.
    sub initArrays
        dim segBmps$(250)           ' Array of all segment bitmap names ("p64" version)
        dim triedBmps(250)          ' Used to track which corridor or room bitmaps tried for new location
        dim miscBitmaps$(1050)      ' Store for unloading purposes of remaining segment bitmaps
        dim levels$(4, (gGridNoRows + 1) * (gGridNoCols + 1))       ' Game has maximum of 3 levels
        dim copyLevels$(4, (gGridNoRows + 1) * (gGridNoCols + 1))   ' Used to temp backup levels

        ' Initialise grid to be equal to the viewable area, stores bitmap name for each grid point
        ' this is the working array that contains one of three levels stored in levels()
        dim grid$(gGridNoRows + 1, gGridNoCols + 1)

        ' Used to reference adjacent cells in grid to current party location
        dim rowOffsets(5)   ' East through North row adjustment for neighboring grid cells
        rowOffsets(1) = 0: rowOffsets(2) = 1: rowOffsets(3) = 0: rowOffsets(4) = -1
        dim colOffsets(5)   ' East through North column adjustment for neighboring grid cells
        colOffsets(1) = 1: colOffsets(2) = 0: colOffsets(3) = -1: colOffsets(4) = 0

        call initArraysOnNew
    end sub

    ' Called each time a new game started
    sub initArraysOnNew
        call initLevelsArray
        call copyLevelToGrid 1
    end sub

    sub initLevelsArray
        for level = 1 to 3
            for row = 1 to gGridNoRows
                for col = 1 to gGridNoCols
                    levels$(level, row * gGridNoCols + col) = gUnvisitedBmp$
                next col
            next row
        next level
    end sub

    ' Tile bitmap naming system
    '   first character "r" is a room tile, "c" is a corridor tile, "u" is unknown
    '   second character is East side "c" corridor exit "d" door exit "w" blank wall
    '   third character is South side "c" corridor exit "d" door exit "w" blank wall
    '   fourth character is West side "c" corridor exit "d" door exit "w" blank wall
    '   fifth character is North side "c" corridor exit "d" door exit "w" blank wall
    '   sixth and seventh character 01 to 11 is feature
    '   eighth character is "p" for version of tile showing current party location or "n" normal
    '   ninth and tenth char bitmap size "64" or "48" or "32"
    sub loadBitmaps
        dim info$(10, 10)
        files DefaultDir$;"\bmp", info$()
        gatewayCode$ = getFeatureNo$("Gateway")
        for file = 1 to val(info$(0, 0))
            bmpName$ = word$(info$(file, 0), 1, ".")
            featureCode$ = extractFeature$(bmpName$)
            tileType$ = extractTileType$(bmpName$)
            tileSize$ = extractSize$(bmpName$)
            tileLocStatus$ = extractLocStatus$(bmpName$)
            loadbmp bmpName$, "bmp\"; info$(file, 0)

            ' Only store party location ("p") 64 bit versions of bitmaps in segBmps$() array,
            ' exclude hellgate and unvisited bitmaps (tileType$ = "u"), Gateway (based on featureCode),
            ' non party location ("n") and non size 64bit from segBmps$() array used to choose new bitmap
            if (featureCode$ <> gatewayCode$) and (tileType$ <>  gUnknown$)_
                and (tileSize$ = "64") and (tileLocStatus$ <> gNormal$) then
                gNoSegBmps = gNoSegBmps + 1
                segBmps$(gNoSegBmps) = bmpName$
            else
                gNoMiscBmps = gNoMiscBmps + 1
                miscBitmaps$(gNoMiscBmps) = bmpName$
            end if
        next file

        ' Load splash image tiles
        files DefaultDir$;"\splash", info$()
        for file = 1 to val(info$(0, 0))
            bmpName$ = "splash"; word$(info$(file, 0), 1, ".")
            loadbmp bmpName$, "splash\"; info$(file, 0)
        next file
    end sub

    ' This sub only called once Hellgate encountered and gHellgateBmpPrefix$ determined
    ' the exits from Hellgate that lead to unvisited tiles are randomnly assigned as
    ' room or corridor
    sub loadHellgateBitmaps
        loadbmp gHellgateBmpPrefix$;"p64", "bmp\uuuuu10p64.bmp"
        loadbmp gHellgateBmpPrefix$;"n64", "bmp\uuuuu10n64.bmp"
        loadbmp gHellgateBmpPrefix$;"p48", "bmp\uuuuu10p48.bmp"
        loadbmp gHellgateBmpPrefix$;"n48", "bmp\uuuuu10n48.bmp"
        loadbmp gHellgateBmpPrefix$;"p32", "bmp\uuuuu10p32.bmp"
        loadbmp gHellgateBmpPrefix$;"n32", "bmp\uuuuu10n32.bmp"
    end sub

    sub openMainWnd
'        nomainwin
        menu #mainWnd, "&File", "&New", handleNew, "&Open...", handleOpen, "&Save", handleSave,_
                       "Save &As...", handleSaveAs
        call setMainWndSize
        open "Citadel of Blood" for window_nf as #mainWnd
        #mainWnd "trapclose closeMainWnd"
        call loadSplash
        call disableGrid
        gGameChanged = 0
'        run "hh.exe help\Citadel of Blood.chm"
        wait
    end sub

    sub setMainWndSize
        WindowWidth = 200
        WindowHeight = 200
        menu #gr, "Dummy"
        open "Getting window size" for graphics_nsb_nf as #gr
        #gr, "home ; down ; posxy penX penY"
        WindowWidth = 64 * gViewCols + 200 - (2 * penX)
        WindowHeight = 64 * gViewRows + 200 - (2 * penY)
        UpperLeftX = DisplayWidth/2 - WindowWidth/2
        UpperLeftY = DisplayHeight/2 - WindowHeight/2
        close #gr
    end sub

    sub closeMainWnd handle$
        ' If current game has changed give user chance to save first
        if gGameChanged then
            confirm "Save changes?"; save$
            if save$ = "yes" then call handleSave
        end if

        confirm "Do you want to quit " + chr$(34) + "Citadel of Blood" + chr$(34); quit$
        if quit$ = "no" then wait
        call unloadBitmaps
        close #handle$
        end
    end sub

    sub unloadBitmaps
        for index = 1 to gNoSegBmps
            bmpName$ = segBmps$(index)
            unloadbmp bmpName$
        next index

        for index = 1 to gNoMiscBmps
            bmpName$ = miscBitmaps$(index)
            unloadbmp bmpName$
        next index

        for row = 1 to 15
            row$ = str$(row)
            if row < 10 then row$ = "0"; str$(row)
            for col = 1 to 25
                col$ = str$(col)
                if col < 10 then col$ = "0"; str$(col)
                bmpName$ = "splash"; row$; col$
                unloadbmp bmpName$
            next col
        next row

        if gHellgateFound then call unloadHellgateBitmaps
    end sub

    sub unloadHellgateBitmaps
        unloadbmp gHellgateBmpPrefix$;"p64"
        unloadbmp gHellgateBmpPrefix$;"n64"
        unloadbmp gHellgateBmpPrefix$;"p48"
        unloadbmp gHellgateBmpPrefix$;"n48"
        unloadbmp gHellgateBmpPrefix$;"p32"
        unloadbmp gHellgateBmpPrefix$;"n32"
    end sub

    ' ------------------------------------------------------------------------------------------------------------
    '
    ' File Menu routines
    '
    ' ------------------------------------------------------------------------------------------------------------

    sub handleNew
        ' If current game has changed give user chance to save first
        if gGameChanged then
            confirm "Save changes?"; save$
            if save$ = "yes" then
                call handleSave
            end if
        end if

        call initGlobalsOnNew
        call initArraysOnNew

        ' Set random version of party Gateway tile in centre of viewport
        grid$(gPartyGridRow,gPartyGridCol) = gCorridor$ ; randomRotate$("cwww"); getFeatureNo$("Gateway");_
                                             gPartyLoc$; "64"
        call disableGrid
        call enableValidGridCells

        call refreshBitmaps
    end sub

    sub handleOpen
        ' If current game has changed give user chance to save first
        if gGameChanged then
            confirm "Save changes?"; save$
            if save$ = "yes" then
                call handleSave
                exit sub
            end if
        end if
    end sub

    sub handleSave
        if gOpenfile$ = "" then
            call handleSaveAs
            exit sub
        end if

        open gOpenfile$ for output as #outFile
        call saveProgress "#outFile"
        close #outFile
    end sub

    sub handleSaveAs
        ' Get file name to save to
        filedialog "Save game progress", "*.prg", fileName$
        if fileName$ = "" then
            exit sub
        end if

        ' Confirm overwrite if it exists already
        if fileExists(DefaultDir$, fileName$) then
            c$ = "This will overwrite existing file contents!" + chr$(13) + chr$(13)
            c$ = c$ + "Continue?"
            confirm c$; answer$
            if answer$ = "no" then exit sub
        end if

        gOpenfile$ = fileName$
        open fileName$ for output as #outFile
        call saveProgress "#outFile"
        close #outFile
    end sub

    sub saveProgress outFile$
    end sub

    ' ------------------------------------------------------------------------------------------------------------
    '
    ' Dungeon grid display routines
    '
    ' ------------------------------------------------------------------------------------------------------------

    sub copyLevelToGrid levelNo
        for row = 1 to gGridNoRows
            for col = 1 to gGridNoCols
                grid$(row, col) = levels$(levelNo, row * gGridNoCols + col)
            next col
        next row
    end sub

    sub copyGridToLevels levelNo
        for row = 1 to gGridNoRows
            for col = 1 to gGridNoCols
                levels$(levelNo, row * gGridNoCols + col) = grid$(row, col)
            next col
        next row
    end sub

    ' When party moves to viewport edge make a copy of levels then resize levels
    ' then restore from copy into levels
    sub resizeLevels resizeE, resizeS, resizeW, resizeN
        ' Backup grid$() array to levels$() array
        call copyGridToLevels gCurrentLevel

        ' Set parameters for restoring to levels() from copyLevels() based on resize direction
        if resizeN then rowInc = gGridRowsInc else rowInc = 0
        if resizeW then colInc = gGridColsInc else colInc = 0

        ' Make copy of current levels() array
        oldGridNoRows = gGridNoRows
        oldGridNoCols = gGridNoCols
        redim copyLevels$(4, (oldGridNoRows + 1) * (oldGridNoCols + 1))
        call backupLevels

        ' Expand levels() array and copy copyLevels() back in
        if resizeS or resizeN then
            gGridNoRows = gGridNoRows + gGridRowsInc
        else
            gGridNoCols = gGridNoCols + gGridColsInc
        end if
        redim levels$(4, (gGridNoRows + 1) * (gGridNoCols + 1))
        call initLevelsArray
        call restoreLevels rowInc, colInc, oldGridNoRows, oldGridNoCols

        ' Expand grid$() array and restore from expanded levels$()
        redim grid$(gGridNoRows + 1, gGridNoCols + 1)
        call copyLevelToGrid gCurrentLevel

        ' Adjust party location
        if resizeN then gPartyGridRow = gPartyGridRow + gGridRowsInc
        if resizeW then gPartyGridCol = gPartyGridCol + gGridColsInc
    end sub

    sub backupLevels
        for level = 1 to 3
            for row = 1 to gGridNoRows
                for col = 1 to gGridNoCols
                    copyLevels$(level, row * gGridNoCols + col) = levels$(level, row * gGridNoCols + col)
                next col
            next row
        next level
    end sub

    sub restoreLevels rowInc, colInc, oldGridNoRows, oldGridNoCols
        for level = 1 to 3
            for row = 1 to oldGridNoRows
                for col = 1 to oldGridNoCols
                    sourceIndex = row * oldGridNoCols + col
                    destIndex = (row + rowInc) * gGridNoCols + (col + colInc)
                    levels$(level, destIndex) = copyLevels$(level, sourceIndex)
                next col
            next row
        next level
    end sub

    sub handleGridClick handle$
        ' Get grid reference for viewport location clicked
        clickedRow = val(mid$(handle$, 14, 2))
        clickedGridRow = gViewRowStart + clickedRow - 1
        clickedCol = val(mid$(handle$, 16, 2))
        clickedGridCol = gViewColStart + clickedCol - 1

        ' Check whether dungeon all deadends situation [6.4] exists prior to Hellgate discovery
        ' and resolve by randomly changing a nearby location bitmap to a valid alternative
        call handleDungeonDeadend clickedGridRow, clickedGridCol

        ' When party chooses to remain on staircase having arrived last turn give them
        ' the option to ascend or descend one level
        call handleStairUse clickedGridRow, clickedGridCol, clickedRow, clickedCol

        ' Change current location bitmap from party location to normal version
        call toggleTileLocStatus$ gPartyGridRow, gPartyGridCol

        ' Get bitmap for new location based on exit type for current location if
        ' party has moved to previously unvisited location
        if grid$(clickedGridRow, clickedGridCol) = gUnvisitedBmp$ then
            grid$(clickedGridRow, clickedGridCol) = getNewLocationBmp$(clickedGridRow, clickedGridCol)
            call handleStairCreation clickedGridRow, clickedGridCol
            call refreshBitmaps
            call handleBitmapRotation$ clickedGridRow, clickedGridCol, clickedRow, clickedCol
            call refreshBitmaps
            call handleMirror grid$(clickedGridRow, clickedGridCol)
        else
            call toggleTileLocStatus$ clickedGridRow, clickedGridCol
            call refreshBitmaps
        end if

        ' Update party location
        gPartyGridRow = clickedGridRow
        gPartyGridCol = clickedGridCol

        ' Check whether grid$() needs extending and/or viewport position adjusting
        call manageGridDimensions clickedGridRow, clickedGridCol
        call manageViewportPos clickedRow, clickedCol
        call refreshBitmaps

        ' Update which grid cells are enabled
        call enableValidGridCells
    end sub

    ' Scan levels$() looking at the tile exit types between adjacent locations looking
    ' for a wildcard beside a door or corridor indicating a deadend situation doesn't exist.
    sub handleDungeonDeadend clickedGridRow, clickedGridCol
        ' Dungeon deadend allowed if Hellgate discovered
        if gHellgateFound then exit sub

        ' Update levels$() array from grid$() array first
        call copyGridToLevels gCurrentLevel

        for level = 1 to 3
            ' Scan rows
            for row = 1 to gGridNoRows
                for col = 1 to gGridNoCols - 1
                    exitPair$ = getExitTypeLevels$(level, row, col, 1); getExitTypeLevels$(level, row, col + 1, 3)
                    if instr("ud uc cu du", exitPair$) then exit sub
                next col
            next row

            ' Scan columns
            for col = 1 to gGridNoCols
                for row = 1 to gGridNoRows - 1
                    exitPair$ = getExitTypeLevels$(level, row, col, 2); getExitTypeLevels$(level, row + 1, col, 4)
                    if instr("ud uc cu du", exitPair$) then exit sub
                next row
            next col
        next level

        ' Spiral out from clicked grid cell looking for bitmap to change on current level
        nextRow = clickedGridRow
        nextCol = clickedGridCol
        repeatCount = 1
        legLen = 1
        legCount = 1
        direction = 1
        do
            call getNextLocation nextRow, nextCol, repeatCount, legLen, legCount, direction
            newBmpName$ = swapInNewBmpName$(nextRow, nextCol)
        loop while newBmpName$ = ""

        if newBmpName$ <> "" then grid$(nextRow, nextCol) = newBmpName$
        call refreshBitmaps
        notice chr$(34); "[6.4] The Citadel maze may not end in a dead end until the Hellgate is located";_
               chr$(34); " A random tile has been swapped to resolve this!"
    end sub

    ' Spiral algorithm goes like this. Take 1 step E, then 1 step S, then 2 steps W then 2 steps N
    ' then 3 steps E then 3 steps S then 4 steps W then 4 steps N etc. (Was tricky to code!)
    sub getNextLocation byref nextRow, byref nextCol, byref repeatCount, byref legLen, byref legCount, byref direction
        select case direction
            case 1
                nextCol = nextCol + 1
            case 2
                nextRow = nextRow + 1
            case 3
                nextCol = nextCol - 1
            case 4
                nextRow = nextRow - 1
        end select
        legCount = legCount + 1
        if legCount > legLen then
            legCount = 1
            direction = direction + 1
            if direction = 5 then direction = 1
            repeatCount = repeatCount + 1
        end if
        if repeatCount > 2 then
            repeatCount = 1
            legLen = legLen + 1
        end if
    end sub

    ' After checking the location has no feature and is in the confines of the grid
    ' change a wall exit to a corridor exit to resolve the dungeon deadend situation
    function swapInNewBmpName$(nextRow, nextCol)
        ' Only change a tile with no feature
        bmpName$ = grid$(nextRow, nextCol)
        if extractFeature$(bmpName$) <> getFeatureNo$("Empty") then exit function

        ' Check tile is in grid$ confines
        if (nextRow < 1) or (nextRow > gGridNoRows) or (nextCol < 1) or (nextCol > gGridNoCols) then exit function

        ' Check a wall is adjacent to unknown exit for the adjacent location and if so change wall exit to door
        exits$ = extractExits$(grid$(nextRow, nextCol))
        requiredExits$ = getRequiredExitsGrid$(nextRow, nextCol)
        for direction = 1 to 4
            if mid$(exits$, direction, 1) = gWall$ and mid$(requiredExits$, direction, 1) = gUnknown$ then
                bmpName$ = left$(bmpName$, direction); gDoor$; right$(bmpName$, 9 - direction)
                swapInNewBmpName$ = bmpName$
                exit function
            end if
        next direction
    end function

    ' When party chooses to remain on staircase having arrived last turn give them
    ' the option to ascend or descend one level
    sub handleStairUse clickedGridRow, clickedGridCol, clickedRow, clickedCol
        bmpName$ = grid$(clickedGridRow, clickedGridCol)
        if extractFeature$(bmpName$) <> getFeatureNo$("Staircase") then exit sub

        ' Can only use stairs if you start turn on them
        if gPartyGridRow <> clickedGridRow or gPartyGridCol <> clickedGridCol then exit sub

        ULx = int((clickedCol + 0.5) * 64)
        ULy = int((clickedRow + 0.5) * 64)
        options$ = ""
        for level = 1 to 3
            level$ = str$(level)
            levelDifference = gCurrentLevel - level
            select case levelDifference
                case 1
                    options$ = options$; "Ascend to Level "; level$; ","
                case 0
                    options$ = options$; "Remain here,"
                case -1
                    options$ = options$; "Descend to Level "; level$; ","
            end select
        next level
        options$ = left$(options$, len(options$) - 1)
        returnValue$ = PopupMenu$(options$, 160, "", "", "", "", ULx, ULy, "Choose Level")
        if not(instr(returnValue$, "Remain")) then
            ' Toggle as leaving staircase on this level
            call toggleTileLocStatus$ clickedGridRow, clickedGridCol

            ' Save grid$() to levels$()
            call copyGridToLevels gCurrentLevel

            ' Copy chosen level in levels$() to grid$() and update current level
            if instr(returnValue$, "Ascend") then
                call copyLevelToGrid gCurrentLevel - 1
                gCurrentLevel = gCurrentLevel - 1
            else
                call copyLevelToGrid gCurrentLevel + 1
                gCurrentLevel = gCurrentLevel + 1
            end if

            ' Need to toggle as have copied in not at location version of tile from levels$()
            call toggleTileLocStatus$ clickedGridRow, clickedGridCol

            call refreshBitmaps
        end if
    end sub

    ' A staircase can only be created if the location is unvisited on all three levels
    ' getNewLocationBmp$() ensures this is the case
    sub handleStairCreation clickedGridRow, clickedGridCol
        bmpName$ = grid$(clickedGridRow, clickedGridCol)
        if extractFeature$(bmpName$) <> getFeatureNo$("Staircase") then exit sub

        bmpName$ = setLocStatus$(bmpName$, gNormal$)
        for level = 1 to 3
            if level <> gCurrentLevel then
                levels$(level, clickedGridRow * gGridNoCols + clickedGridCol) = bmpName$
            end if
        next level
    end sub

    ' Mirrors described in [13.7]
    sub handleMirror bmpName$
        ' Mirrors do nothing if Hellgate already located
        if gHellgateFound or extractFeature$(bmpName$) <> getFeatureNo$("Mirror") then exit sub

        ' If first mirror then roll on table 13.9 to see what level Hellgate is on
        if not(gFirstMirrorFound) then
            roll = getD6()
            select case roll
                case 1
                    gHellgateLevel = 1
                case 2, 3
                    gHellgateLevel = 2
                case 4, 5, 6
                    gHellgateLevel = 3
            end select

            ' If on Hellgate level roll on table 13.9 to see how many new segments to visit
            ' before locating it
            if gCurrentLevel = gHellgateLevel then
                roll = getD6()
                gHellgateSegsAway = roll + 2
            end if
            gFirstMirrorFound = 1

            ' Advise user
            if gCurrentLevel = gHellgateLevel then
                notice "Mirror reveals the Hellgate is on this level and ";_
                    gHellgateSegsAway; " unvisited segments away [13.7]"
            else
                notice "Mirror reveals the Hellgate is on level "; gHellgateLevel;_
                    ", visit a mirror on that level to see how close it is [13.7]"
            end if
        else
            ' Having now arrived on Hellgate level and encountered a Mirror roll to see
            ' how many new segments to visit before encountering it
            if gCurrentLevel = gHellgateLevel and not(gSecondMirrorFound) then
                roll = getD6()
                gHellgateSegsAway = roll + 2
                notice "Mirror reveals the Hellgate is "; gHellgateSegsAway; " unvisited segments away [13.7]"
                gSecondMirrorFound = 1
            end if
        end if
    end sub

    sub handleBitmapRotation$ clickedGridRow, clickedGridCol, clickedRow, clickedCol
        ' No rotation of Hellgate since non fixed exits randomly set anyway
        if grid$(clickedGridRow, clickedGridCol) = getHellgateBmpName$ then exit sub

        ' Rotate new bitmap exit string three times to see if alternate tile rotations are valid
        requiredExits$ = getRequiredExitsGrid$(clickedGridRow, clickedGridCol)
        exits$ = extractExits$(grid$(clickedGridRow, clickedGridCol))
        validRotations$ = exits$
        validRotCount = 1
        for rotationCount = 1 to 3
            exits$ = rotate$(exits$)
            ' Exclude non unique rotations
            if not(instr(validRotations$, exits$)) and validBitmap(requiredExits$, exits$) then
                validRotations$ = validRotations$; " "; exits$
                validRotCount = validRotCount + 1
            end if
        next rotationCount
        ULx = int((clickedCol + 0.5) * 64)
        ULy = int((clickedRow + 0.5) * 64)
        options$ = "Rotate Clockwise,Leave as is"
        if validRotCount > 1 then
            do
                returnValue$ = PopupMenu$(options$, 160, "", "", "", "", ULx, ULy, "Choose Rotation")
                if returnValue$ = "Rotate Clockwise" then
                    bmpName$ = grid$(clickedGridRow, clickedGridCol)
                    validRotations$ = rotateWordsString$(validRotations$, " ")
                    exits$ = word$(validRotations$, 1)
                    grid$(clickedGridRow, clickedGridCol) = setExits$(bmpName$, exits$)
                    call refreshBitmaps
                end if
            loop until returnValue$ = "Leave as is"
        end if
    end sub

    sub manageGridDimensions clickedGridRow, clickedGridCol
        if clickedGridRow = 2 then call resizeLevels 0, 0, 0, 1
        if clickedGridRow = gGridNoRows - 1 then call resizeLevels 0, 1, 0, 0
        if clickedGridCol = 2 then call resizeLevels 0, 0, 1, 0
        if clickedGridCol = gGridNoCols - 1 then call resizeLevels 1, 0, 0, 0
    end sub

    sub manageViewportPos clickedRow, clickedCol
        if clickedRow = 1 then gViewRowStart = gViewRowStart - gGridRowsInc
        if clickedRow = gViewRows then gViewRowStart = gViewRowStart + gGridRowsInc
        if clickedCol = 1 then gViewColStart = gViewColStart - gGridColsInc
        if clickedCol = gViewCols then gViewColStart = gViewColStart + gGridColsInc
    end sub

    sub refreshBitmaps
        for viewRow = 1 to gViewRows
            for viewCol = 1 to gViewCols
                buttonhandle$ = getCellNameFromViewRef$(viewRow, viewCol)
                gridRow = gViewRowStart + viewRow - 1
                gridCol = gViewColStart + viewCol - 1
                bmpName$ = grid$(gridRow, gridCol)
                #buttonhandle$ "bitmap "; bmpName$
            next viewCol
        next viewRow
    end sub

    sub loadSplash
        for viewRow = 1 to gViewRows
            for viewCol = 1 to gViewCols
                buttonhandle$ = getCellNameFromViewRef$(viewRow, viewCol)
                bmpName$ = "splash"; pad$(viewRow); pad$(viewCol)
                #buttonhandle$ "bitmap "; bmpName$
            next viewCol
        next viewRow
    end sub

    function getCellNameFromViewRef$(viewRow, viewCol)
        getCellNameFromViewRef$ = "#mainWnd.cell"; pad$(viewRow); pad$(viewCol)
    end function

    function getCellNameFromGridRef$(gridRow, gridCol)
        viewRow = gridRow - gViewRowStart + 1
        viewCol = gridCol - gViewColStart + 1
        getCellNameFromGridRef$ = "#mainWnd.cell"; pad$(viewRow); pad$(viewCol)
    end function

    sub disableGrid
        for viewRow = 1 to gViewRows
            for viewCol = 1 to gViewCols
                buttonhandle$ = getCellNameFromViewRef$(viewRow, viewCol)
                #buttonhandle$ "disable"
            next viewCol
        next viewRow
    end sub

    ' Enable adjacent grid cells to current party location that can be reached via door or corridor
    sub enableValidGridCells
        call disableGrid
        for dirIndex = 1 to 4
            exitType$ = getExitTypeGrid$(gPartyGridRow, gPartyGridCol, dirIndex)
            if exitType$ <> gWall$ then
                buttonhandle$ = getCellNameFromGridRef$(gPartyGridRow + rowOffsets(dirIndex),_
                                                        gPartyGridCol + colOffsets(dirIndex))
                #buttonhandle$ "enable"
            end if
        next dirIndex

        ' Enable current location if it's a staircase
        bmpName$ = grid$(gPartyGridRow, gPartyGridCol)
        if extractFeature$(bmpName$) = getFeatureNo$("Staircase") then
            buttonhandle$ = getCellNameFromGridRef$(gPartyGridRow, gPartyGridCol)
            #buttonhandle$ "enable"
        end if
    end sub

    ' ------------------------------------------------------------------------------------------------------------
    '
    ' Party movement routines
    '
    ' ------------------------------------------------------------------------------------------------------------

    function getNewLocationBmp$(newGridRow, newGridCol)

options$ = "ccwww09p64,cwcww09p64,cwwcw09p64,cwwwc09p64,rdddd06p64,skip"
bmpName$ = PopupMenu$(options$, 160, "", "", "", "", 50, 50, "Choose Bitmap")
if bmpName$ <> "skip" then
    getNewLocationBmp$ = bmpName$
    exit function
end if

        validBmp = 0

        ' Determine whether party has moved through door or corridor from current location
        ' movedDir 1=East, 2=South etc
        movedDir = getAdjGridsDir(gPartyGridRow, gPartyGridCol, newGridRow, newGridCol)
        exitType$ = getExitTypeGrid$(gPartyGridRow, gPartyGridCol, movedDir)

        ' Determine whether new location is room or corridor, moving through door exit
        ' enters room 80% of the time, otherwise move is into a corridor
        if exitType$ = gDoor$ and (rnd(1) > 0.2) then
            tileType$ = gRoom$
        else
            tileType$ = gCorridor$
        end if

        ' Get required string to match new location surrounding tiles
        requiredExits$ = getRequiredExitsGrid$(newGridRow, newGridCol)

        ' Probability of Hellgate is 1 in 200 (based on counter sheets) alternatively
        ' if mirrors have revealed location [13.7]
        randHellgate = int(rnd(1) * 200 + 1)
        if ((randHellgate = 1) and not(gHellgateFound)) or_
            (gCurrentLevel = gHellgateLevel and gHellgateSegsAway = 1 and not(gHellgateFound)) then
                getNewLocationBmp$ = getHellgateBmpName$(requiredExits$)
                call loadHellgateBitmaps
                gHellgateFound = 1
                exit function
        end if
        ' Count down to discovery of Hellgate
        if not(gHellgateFound) and gCurrentLevel = gHellgateLevel and gHellgateSegsAway > 0 then
            gHellgateSegsAway = gHellgateSegsAway - 1
        end if

[searchSegments]
        ' Randomnly try all segment bitmaps and stop when one found, with additional
        ' counters added from those that came with board game, I "think" there will always
        ' be one that fits!
        redim triedBmps(gNoSegBmps + 1)
        noUntried = gNoSegBmps
        for index = 1 to gNoSegBmps
            continue = 1
            segIndex = getRandomIndex(gNoSegBmps, noUntried)

            ' Check it's room or corridor as required
            if extractTileType$(segBmps$(segIndex)) <> tileType$ then continue = 0

            ' Check exits are valid
            exits$ = extractExits$(segBmps$(segIndex))
            if continue and validBitmap(requiredExits$, exits$) then
                ' If room then override random equal probability feature with preset feature probabilities
                if tileType$ = gRoom$ then
                    do
                        getNewLocationBmp$ = setRanRoomFeature$(segBmps$(segIndex))
                    loop until checkStairValid(newGridRow, newGridCol, getNewLocationBmp$)
                else
                    getNewLocationBmp$ = segBmps$(segIndex)
                end if
                validBmp = 1
                exit for
            end if
        next index

        ' If unable to find valid room bitmap try for a corridor instead
        if not(validBmp) and tileType$ = gRoom$ then
            tileType$ = gCorridor$
            goto [searchSegments]
        end if

        if trim$(getNewLocationBmp$) = "" then
            getNewLocationBmp$ = gUnvisitedBmp$
            notice "This shouldn't happen requiredExits$ "; requiredExits$; " exits$ "; exits$
        end if
    end function

    ' Get direction to move from first location to adjacent second location 1=East 2=South etc
    function getAdjGridsDir(firstRow, firstCol, secRow, secCol)
        if firstRow = secRow then
            if secCol > firstCol then getAdjGridsDir = 1 else getAdjGridsDir = 3
        else
            if secRow > firstRow then getAdjGridsDir = 2 else getAdjGridsDir = 4
        end if
    end function

    ' Gets a random index using triedBmps() redimensioned for this purpose as required
    function getRandomIndex(noElements, byref noUntried)
        randCount = int(rnd(1) * noUntried + 1)
        for index = 1 to noElements
            if triedBmps(index) = 0 then randCount = randCount - 1
            if randCount = 0 then
                getRandomIndex = index
                triedBmps(index) = 1
                noUntried = noUntried - 1
                exit for
            end if
        next index
    end function

    ' For locations adjacent to hellgate that haven't been visited yet by party assign random
    ' exit from hellgate location of corridor or door
    function getHellgateBmpName$(requiredExits$)
        for direction = 1 to 4
            if mid$(requiredExits$, direction, 1) = gUnknown$ then
                if rnd(1) > 0.5 then
                    exitType$ = gCorridor$
                else
                    exitType$ = gDoor$
                end if
                requiredExits$ = left$(requiredExits$, direction - 1); exitType$;_
                                 right$(requiredExits$, 4 - direction)
            end if
        next direction
        gHellgateBmpPrefix$ = gRoom$; requiredExits$; getFeatureNo$("Hellgate")
        getHellgateBmpName$ = gHellgateBmpPrefix$; gPartyLoc$; "64"
    end function

    ' For each direction in requiredExits$ compares with candidateExits$, if all directions
    ' match the new location bitmap will fit with adjacent tiles (gUnknown$ acts as wild character)
    function validBitmap(requiredExits$, candidateExits$)
        validBitmap = 1
        for index = 1 to 4
            reqExit$ = mid$(requiredExits$, index, 1)
            candExit$ = mid$(candidateExits$, index, 1)
            if (reqExit$ <> candExit$) and (reqExit$ <> gUnknown$) then
                validBitmap = 0
            end if
        next index
    end function

    function setRanRoomFeature$(bmpName$)
        ' 37% no feature based on counter sheets
        randIsFeature = rnd(1)
        if randIsFeature <= .50 then
            setRanRoomFeature$ = setFeature$(bmpName$, "Empty")
        else
            ' Staircase and Mirror 20% each, remaining 6 of first 8 in gFeatures$ 10% each based on counter sheets
            randIsFeature = int(rnd(1) * 8 + 1)
            if (gSecondMirrorFound and word$(gFeatures$, randIsFeature) = "Mirror") or gHellgateFound then
                do
                    randIsFeature = int(rnd(1) * 8 + 1)
                loop while word$(gFeatures$, randIsFeature) = "Mirror"
            end if
            select case randIsFeature
'                case 5, 9
'                    setRanRoomFeature$ = setFeature$(bmpName$, "Mirror")
'                case 6, 10
'                    setRanRoomFeature$ = setFeature$(bmpName$, "Staircase")
                case else
                    setRanRoomFeature$ = setFeature$(bmpName$, word$(gFeatures$, randIsFeature))
            end select
        end if
    end function

    ' A staircase can only be created if the location is unvisited on all three levels and
    ' on each level it fits adjacent tiles regarding corresponding exit types
    function checkStairValid(newGridRow, newGridCol, newBmpName$)
        checkStairValid = 1
        if extractFeature$(newBmpName$) = getFeatureNo$("Staircase") then
            for level = 1 to 3
                if level <> gCurrentLevel then
                    if levels$(level, newGridRow * gGridNoCols + newGridCol) <> gUnvisitedBmp$ then
                        checkStairValid = 0
                    else
                        ' Check newBmpName$ will fit any adjacent tiles on the level
                        requiredExits$ = getRequiredExitsLevels$(level, newGridRow, newGridCol)
                        exits$ = extractExits$(newBmpName$)
                        if not(validBitmap(requiredExits$, exits$)) then checkStairValid = 0
                    end if
                end if
            next level
        end if
    end function

    ' ------------------------------------------------------------------------------------------------------------
    '
    ' Bitmap manipulation routines
    '
    ' ------------------------------------------------------------------------------------------------------------

    function extractTileType$(bmpName$)
        extractTileType$ = left$(bmpName$, 1)
    end function

    function extractExits$(bmpName$)
        extractExits$ = mid$(bmpName$, 2, 4)
    end function

    function extractFeature$(bmpName$)
        extractFeature$ = mid$(bmpName$, 6, 2)
    end function

    function extractLocStatus$(bmpName$)
        extractLocStatus$ = mid$(bmpName$, 8, 1)
    end function

    function extractSize$(bmpName$)
        extractSize$ = right$(bmpName$, 2)
    end function

    ' Get exit type (w=wall, d=door or c=corridor) for a grid location in given direction 1=East 2=South etc
    function getExitTypeGrid$(row, col, direction)
        bmpName$ = grid$(row, col)
        exits$ = extractExits$(bmpName$)
        getExitTypeGrid$ = mid$(exits$, direction, 1)
    end function

    ' Get exit type (w=wall, d=door or c=corridor) for a grid location in given direction 1=East 2=South etc
    function getExitTypeLevels$(level, row, col, direction)
        bmpName$ = levels$(level, row * gGridNoCols + col)
        exits$ = extractExits$(bmpName$)
        getExitTypeLevels$ = mid$(exits$, direction, 1)
    end function

    ' Finds position of feature$ in gFeatures$ and returns it padded to two character string
    function getFeatureNo$(feature$)
        select case feature$
            case "Altar"
                getFeatureNo$ = "01"
            case "Artwork"
                getFeatureNo$ = "02"
            case "Fountain"
                getFeatureNo$ = "03"
            case "Furniture"
                getFeatureNo$ = "04"
            case "Mirror"
                getFeatureNo$ = "05"
            case "Staircase"
                getFeatureNo$ = "06"
            case "Statue"
                getFeatureNo$ = "07"
            case "Trap"
                getFeatureNo$ = "08"
            case "Empty"
                getFeatureNo$ = "09"
            case "Hellgate"
                getFeatureNo$ = "10"
            case "Gateway"
                getFeatureNo$ = "11"
        end select
    end function

    ' For passed grid location returns string for the four adjacent location exit types into
    ' the passed grid location
    function getRequiredExitsGrid$(row, col)
        for dirOut = 1 to 4
            adjRow = row + rowOffsets(dirOut)
            adjCol = col + colOffsets(dirOut)
            exit$ = getExitTypeGrid$(adjRow, adjCol, getOppositeDir(dirOut))
            getRequiredExitsGrid$ = getRequiredExitsGrid$; exit$
        next dirOut
    end function

    function getRequiredExitsLevels$(level, row, col)
        for dirOut = 1 to 4
            adjRow = row + rowOffsets(dirOut)
            adjCol = col + colOffsets(dirOut)
            exit$ = getExitTypeLevels$(level, adjRow, adjCol, getOppositeDir(dirOut))
            getRequiredExitsLevels$ = getRequiredExitsLevels$; exit$
        next dirOut
    end function

    function pad$(featureNo)
        pad$ = str$(featureNo)
        if featureNo < 10 then pad$ = "0";pad$
    end function

    function setExits$(bmpName$, exits$)
        setExits$ = left$(bmpName$, 1); exits$; right$(bmpName$, 5)
    end function

    function setFeature$(bmpName$, feature$)
        setFeature$ = left$(bmpName$, 5); getFeatureNo$(feature$); right$(bmpName$, 3)
    end function

    function setLocStatus$(bmpName$, locStatus$)
        setLocStatus$ = left$(bmpName$, 7); locStatus$; right$(bmpName$, 2)
    end function

    sub toggleTileLocStatus$ clickedGridRow, clickedGridCol
        bmpName$ = grid$(clickedGridRow, clickedGridCol)
        if mid$(bmpName$, 8, 1) = gNormal$ then
            grid$(clickedGridRow, clickedGridCol) = left$(bmpName$, 7); gPartyLoc$; right$(bmpName$, 2)
        else
            grid$(clickedGridRow, clickedGridCol) = left$(bmpName$, 7); gNormal$; right$(bmpName$, 2)
        end if
    end sub

    ' ------------------------------------------------------------------------------------------------------------
    '
    ' General utility routines
    '
    ' ------------------------------------------------------------------------------------------------------------

    function getD6()
        getD6 = int(rnd(1) * 6) + 1
    end function

    ' Switches East with West and North with South
    function getOppositeDir(direction)
        getOppositeDir = (direction + 2)
        if getOppositeDir > 4 then getOppositeDir = getOppositeDir - 4
    end function

    ' Input is string of characters output is that string character sequence randomly rotated
    ' e.g. a string of "dwww" shows a door exit East and three blank walls, output would be one of
    ' "dwww", "wdww", wwdw", or "wwwd"
    function randomRotate$(string$)
        lenStr = len(string$)
        randNo = int(rnd(1) * lenStr) + 1
        for rotate = 1 to randNo
            string$ = right$(string$, 1) + left$(string$, lenStr - 1)
        next rotate
        randomRotate$ = string$
    end function

    ' Input is string of characters output is that string character rotated once
    ' e.g. a string of "dwww" shows a door exit East and three blank walls, output would be "wdww"
    function rotate$(string$)
        lenStr = len(string$)
        rotate$ = right$(string$, 1); left$(string$, lenStr - 1)
    end function

    ' Takes first word in list of words and places it at the end
    function rotateWordsString$(string$, separator$)
        trimmed$ = trim$(string$)
        dim words$(100)
        do
            count = count + 1
            words$(count) = word$(trimmed$, count, separator$)
        loop while words$(count) <> ""
        for pos = 2 to count - 1
            rotateWordsString$ = rotateWordsString$; words$(pos); separator$
        next pos
        rotateWordsString$ = rotateWordsString$; words$(1)
    end function

    ' ------------------------------------------------------------------------------------------------------------
    '
    ' Create Buttons
    '
    ' ------------------------------------------------------------------------------------------------------------

    sub createButtons
        bmpbutton #mainWnd.cell0101, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 0, 0
        bmpbutton #mainWnd.cell0102, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 64, 0
        bmpbutton #mainWnd.cell0103, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 128, 0
        bmpbutton #mainWnd.cell0104, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 192, 0
        bmpbutton #mainWnd.cell0105, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 256, 0
        bmpbutton #mainWnd.cell0106, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 320, 0
        bmpbutton #mainWnd.cell0107, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 384, 0
        bmpbutton #mainWnd.cell0108, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 448, 0
        bmpbutton #mainWnd.cell0109, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 512, 0
        bmpbutton #mainWnd.cell0110, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 576, 0
        bmpbutton #mainWnd.cell0111, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 640, 0
        bmpbutton #mainWnd.cell0112, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 704, 0
        bmpbutton #mainWnd.cell0113, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 768, 0
        bmpbutton #mainWnd.cell0114, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 832, 0
        bmpbutton #mainWnd.cell0115, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 896, 0
        bmpbutton #mainWnd.cell0116, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 960, 0
        bmpbutton #mainWnd.cell0117, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1024, 0
        bmpbutton #mainWnd.cell0118, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1088, 0
        bmpbutton #mainWnd.cell0119, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1152, 0
        bmpbutton #mainWnd.cell0120, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1216, 0
        bmpbutton #mainWnd.cell0121, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1280, 0
        bmpbutton #mainWnd.cell0122, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1344, 0
        bmpbutton #mainWnd.cell0123, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1408, 0
        bmpbutton #mainWnd.cell0124, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1472, 0
        bmpbutton #mainWnd.cell0125, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1536, 0
        bmpbutton #mainWnd.cell0201, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 0, 64
        bmpbutton #mainWnd.cell0202, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 64, 64
        bmpbutton #mainWnd.cell0203, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 128, 64
        bmpbutton #mainWnd.cell0204, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 192, 64
        bmpbutton #mainWnd.cell0205, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 256, 64
        bmpbutton #mainWnd.cell0206, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 320, 64
        bmpbutton #mainWnd.cell0207, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 384, 64
        bmpbutton #mainWnd.cell0208, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 448, 64
        bmpbutton #mainWnd.cell0209, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 512, 64
        bmpbutton #mainWnd.cell0210, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 576, 64
        bmpbutton #mainWnd.cell0211, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 640, 64
        bmpbutton #mainWnd.cell0212, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 704, 64
        bmpbutton #mainWnd.cell0213, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 768, 64
        bmpbutton #mainWnd.cell0214, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 832, 64
        bmpbutton #mainWnd.cell0215, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 896, 64
        bmpbutton #mainWnd.cell0216, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 960, 64
        bmpbutton #mainWnd.cell0217, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1024, 64
        bmpbutton #mainWnd.cell0218, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1088, 64
        bmpbutton #mainWnd.cell0219, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1152, 64
        bmpbutton #mainWnd.cell0220, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1216, 64
        bmpbutton #mainWnd.cell0221, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1280, 64
        bmpbutton #mainWnd.cell0222, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1344, 64
        bmpbutton #mainWnd.cell0223, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1408, 64
        bmpbutton #mainWnd.cell0224, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1472, 64
        bmpbutton #mainWnd.cell0225, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1536, 64
        bmpbutton #mainWnd.cell0301, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 0, 128
        bmpbutton #mainWnd.cell0302, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 64, 128
        bmpbutton #mainWnd.cell0303, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 128, 128
        bmpbutton #mainWnd.cell0304, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 192, 128
        bmpbutton #mainWnd.cell0305, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 256, 128
        bmpbutton #mainWnd.cell0306, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 320, 128
        bmpbutton #mainWnd.cell0307, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 384, 128
        bmpbutton #mainWnd.cell0308, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 448, 128
        bmpbutton #mainWnd.cell0309, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 512, 128
        bmpbutton #mainWnd.cell0310, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 576, 128
        bmpbutton #mainWnd.cell0311, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 640, 128
        bmpbutton #mainWnd.cell0312, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 704, 128
        bmpbutton #mainWnd.cell0313, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 768, 128
        bmpbutton #mainWnd.cell0314, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 832, 128
        bmpbutton #mainWnd.cell0315, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 896, 128
        bmpbutton #mainWnd.cell0316, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 960, 128
        bmpbutton #mainWnd.cell0317, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1024, 128
        bmpbutton #mainWnd.cell0318, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1088, 128
        bmpbutton #mainWnd.cell0319, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1152, 128
        bmpbutton #mainWnd.cell0320, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1216, 128
        bmpbutton #mainWnd.cell0321, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1280, 128
        bmpbutton #mainWnd.cell0322, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1344, 128
        bmpbutton #mainWnd.cell0323, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1408, 128
        bmpbutton #mainWnd.cell0324, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1472, 128
        bmpbutton #mainWnd.cell0325, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1536, 128
        bmpbutton #mainWnd.cell0401, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 0, 192
        bmpbutton #mainWnd.cell0402, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 64, 192
        bmpbutton #mainWnd.cell0403, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 128, 192
        bmpbutton #mainWnd.cell0404, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 192, 192
        bmpbutton #mainWnd.cell0405, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 256, 192
        bmpbutton #mainWnd.cell0406, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 320, 192
        bmpbutton #mainWnd.cell0407, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 384, 192
        bmpbutton #mainWnd.cell0408, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 448, 192
        bmpbutton #mainWnd.cell0409, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 512, 192
        bmpbutton #mainWnd.cell0410, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 576, 192
        bmpbutton #mainWnd.cell0411, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 640, 192
        bmpbutton #mainWnd.cell0412, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 704, 192
        bmpbutton #mainWnd.cell0413, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 768, 192
        bmpbutton #mainWnd.cell0414, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 832, 192
        bmpbutton #mainWnd.cell0415, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 896, 192
        bmpbutton #mainWnd.cell0416, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 960, 192
        bmpbutton #mainWnd.cell0417, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1024, 192
        bmpbutton #mainWnd.cell0418, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1088, 192
        bmpbutton #mainWnd.cell0419, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1152, 192
        bmpbutton #mainWnd.cell0420, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1216, 192
        bmpbutton #mainWnd.cell0421, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1280, 192
        bmpbutton #mainWnd.cell0422, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1344, 192
        bmpbutton #mainWnd.cell0423, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1408, 192
        bmpbutton #mainWnd.cell0424, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1472, 192
        bmpbutton #mainWnd.cell0425, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1536, 192
        bmpbutton #mainWnd.cell0501, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 0, 256
        bmpbutton #mainWnd.cell0502, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 64, 256
        bmpbutton #mainWnd.cell0503, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 128, 256
        bmpbutton #mainWnd.cell0504, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 192, 256
        bmpbutton #mainWnd.cell0505, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 256, 256
        bmpbutton #mainWnd.cell0506, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 320, 256
        bmpbutton #mainWnd.cell0507, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 384, 256
        bmpbutton #mainWnd.cell0508, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 448, 256
        bmpbutton #mainWnd.cell0509, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 512, 256
        bmpbutton #mainWnd.cell0510, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 576, 256
        bmpbutton #mainWnd.cell0511, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 640, 256
        bmpbutton #mainWnd.cell0512, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 704, 256
        bmpbutton #mainWnd.cell0513, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 768, 256
        bmpbutton #mainWnd.cell0514, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 832, 256
        bmpbutton #mainWnd.cell0515, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 896, 256
        bmpbutton #mainWnd.cell0516, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 960, 256
        bmpbutton #mainWnd.cell0517, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1024, 256
        bmpbutton #mainWnd.cell0518, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1088, 256
        bmpbutton #mainWnd.cell0519, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1152, 256
        bmpbutton #mainWnd.cell0520, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1216, 256
        bmpbutton #mainWnd.cell0521, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1280, 256
        bmpbutton #mainWnd.cell0522, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1344, 256
        bmpbutton #mainWnd.cell0523, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1408, 256
        bmpbutton #mainWnd.cell0524, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1472, 256
        bmpbutton #mainWnd.cell0525, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1536, 256
        bmpbutton #mainWnd.cell0601, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 0, 320
        bmpbutton #mainWnd.cell0602, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 64, 320
        bmpbutton #mainWnd.cell0603, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 128, 320
        bmpbutton #mainWnd.cell0604, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 192, 320
        bmpbutton #mainWnd.cell0605, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 256, 320
        bmpbutton #mainWnd.cell0606, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 320, 320
        bmpbutton #mainWnd.cell0607, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 384, 320
        bmpbutton #mainWnd.cell0608, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 448, 320
        bmpbutton #mainWnd.cell0609, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 512, 320
        bmpbutton #mainWnd.cell0610, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 576, 320
        bmpbutton #mainWnd.cell0611, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 640, 320
        bmpbutton #mainWnd.cell0612, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 704, 320
        bmpbutton #mainWnd.cell0613, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 768, 320
        bmpbutton #mainWnd.cell0614, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 832, 320
        bmpbutton #mainWnd.cell0615, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 896, 320
        bmpbutton #mainWnd.cell0616, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 960, 320
        bmpbutton #mainWnd.cell0617, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1024, 320
        bmpbutton #mainWnd.cell0618, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1088, 320
        bmpbutton #mainWnd.cell0619, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1152, 320
        bmpbutton #mainWnd.cell0620, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1216, 320
        bmpbutton #mainWnd.cell0621, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1280, 320
        bmpbutton #mainWnd.cell0622, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1344, 320
        bmpbutton #mainWnd.cell0623, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1408, 320
        bmpbutton #mainWnd.cell0624, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1472, 320
        bmpbutton #mainWnd.cell0625, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1536, 320
        bmpbutton #mainWnd.cell0701, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 0, 384
        bmpbutton #mainWnd.cell0702, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 64, 384
        bmpbutton #mainWnd.cell0703, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 128, 384
        bmpbutton #mainWnd.cell0704, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 192, 384
        bmpbutton #mainWnd.cell0705, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 256, 384
        bmpbutton #mainWnd.cell0706, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 320, 384
        bmpbutton #mainWnd.cell0707, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 384, 384
        bmpbutton #mainWnd.cell0708, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 448, 384
        bmpbutton #mainWnd.cell0709, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 512, 384
        bmpbutton #mainWnd.cell0710, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 576, 384
        bmpbutton #mainWnd.cell0711, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 640, 384
        bmpbutton #mainWnd.cell0712, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 704, 384
        bmpbutton #mainWnd.cell0713, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 768, 384
        bmpbutton #mainWnd.cell0714, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 832, 384
        bmpbutton #mainWnd.cell0715, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 896, 384
        bmpbutton #mainWnd.cell0716, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 960, 384
        bmpbutton #mainWnd.cell0717, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1024, 384
        bmpbutton #mainWnd.cell0718, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1088, 384
        bmpbutton #mainWnd.cell0719, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1152, 384
        bmpbutton #mainWnd.cell0720, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1216, 384
        bmpbutton #mainWnd.cell0721, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1280, 384
        bmpbutton #mainWnd.cell0722, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1344, 384
        bmpbutton #mainWnd.cell0723, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1408, 384
        bmpbutton #mainWnd.cell0724, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1472, 384
        bmpbutton #mainWnd.cell0725, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1536, 384
        bmpbutton #mainWnd.cell0801, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 0, 448
        bmpbutton #mainWnd.cell0802, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 64, 448
        bmpbutton #mainWnd.cell0803, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 128, 448
        bmpbutton #mainWnd.cell0804, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 192, 448
        bmpbutton #mainWnd.cell0805, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 256, 448
        bmpbutton #mainWnd.cell0806, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 320, 448
        bmpbutton #mainWnd.cell0807, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 384, 448
        bmpbutton #mainWnd.cell0808, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 448, 448
        bmpbutton #mainWnd.cell0809, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 512, 448
        bmpbutton #mainWnd.cell0810, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 576, 448
        bmpbutton #mainWnd.cell0811, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 640, 448
        bmpbutton #mainWnd.cell0812, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 704, 448
        bmpbutton #mainWnd.cell0813, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 768, 448
        bmpbutton #mainWnd.cell0814, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 832, 448
        bmpbutton #mainWnd.cell0815, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 896, 448
        bmpbutton #mainWnd.cell0816, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 960, 448
        bmpbutton #mainWnd.cell0817, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1024, 448
        bmpbutton #mainWnd.cell0818, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1088, 448
        bmpbutton #mainWnd.cell0819, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1152, 448
        bmpbutton #mainWnd.cell0820, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1216, 448
        bmpbutton #mainWnd.cell0821, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1280, 448
        bmpbutton #mainWnd.cell0822, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1344, 448
        bmpbutton #mainWnd.cell0823, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1408, 448
        bmpbutton #mainWnd.cell0824, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1472, 448
        bmpbutton #mainWnd.cell0825, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1536, 448
        bmpbutton #mainWnd.cell0901, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 0, 512
        bmpbutton #mainWnd.cell0902, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 64, 512
        bmpbutton #mainWnd.cell0903, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 128, 512
        bmpbutton #mainWnd.cell0904, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 192, 512
        bmpbutton #mainWnd.cell0905, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 256, 512
        bmpbutton #mainWnd.cell0906, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 320, 512
        bmpbutton #mainWnd.cell0907, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 384, 512
        bmpbutton #mainWnd.cell0908, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 448, 512
        bmpbutton #mainWnd.cell0909, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 512, 512
        bmpbutton #mainWnd.cell0910, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 576, 512
        bmpbutton #mainWnd.cell0911, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 640, 512
        bmpbutton #mainWnd.cell0912, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 704, 512
        bmpbutton #mainWnd.cell0913, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 768, 512
        bmpbutton #mainWnd.cell0914, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 832, 512
        bmpbutton #mainWnd.cell0915, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 896, 512
        bmpbutton #mainWnd.cell0916, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 960, 512
        bmpbutton #mainWnd.cell0917, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1024, 512
        bmpbutton #mainWnd.cell0918, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1088, 512
        bmpbutton #mainWnd.cell0919, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1152, 512
        bmpbutton #mainWnd.cell0920, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1216, 512
        bmpbutton #mainWnd.cell0921, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1280, 512
        bmpbutton #mainWnd.cell0922, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1344, 512
        bmpbutton #mainWnd.cell0923, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1408, 512
        bmpbutton #mainWnd.cell0924, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1472, 512
        bmpbutton #mainWnd.cell0925, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1536, 512
        bmpbutton #mainWnd.cell1001, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 0, 576
        bmpbutton #mainWnd.cell1002, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 64, 576
        bmpbutton #mainWnd.cell1003, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 128, 576
        bmpbutton #mainWnd.cell1004, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 192, 576
        bmpbutton #mainWnd.cell1005, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 256, 576
        bmpbutton #mainWnd.cell1006, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 320, 576
        bmpbutton #mainWnd.cell1007, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 384, 576
        bmpbutton #mainWnd.cell1008, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 448, 576
        bmpbutton #mainWnd.cell1009, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 512, 576
        bmpbutton #mainWnd.cell1010, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 576, 576
        bmpbutton #mainWnd.cell1011, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 640, 576
        bmpbutton #mainWnd.cell1012, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 704, 576
        bmpbutton #mainWnd.cell1013, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 768, 576
        bmpbutton #mainWnd.cell1014, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 832, 576
        bmpbutton #mainWnd.cell1015, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 896, 576
        bmpbutton #mainWnd.cell1016, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 960, 576
        bmpbutton #mainWnd.cell1017, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1024, 576
        bmpbutton #mainWnd.cell1018, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1088, 576
        bmpbutton #mainWnd.cell1019, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1152, 576
        bmpbutton #mainWnd.cell1020, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1216, 576
        bmpbutton #mainWnd.cell1021, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1280, 576
        bmpbutton #mainWnd.cell1022, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1344, 576
        bmpbutton #mainWnd.cell1023, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1408, 576
        bmpbutton #mainWnd.cell1024, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1472, 576
        bmpbutton #mainWnd.cell1025, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1536, 576
        bmpbutton #mainWnd.cell1101, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 0, 640
        bmpbutton #mainWnd.cell1102, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 64, 640
        bmpbutton #mainWnd.cell1103, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 128, 640
        bmpbutton #mainWnd.cell1104, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 192, 640
        bmpbutton #mainWnd.cell1105, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 256, 640
        bmpbutton #mainWnd.cell1106, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 320, 640
        bmpbutton #mainWnd.cell1107, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 384, 640
        bmpbutton #mainWnd.cell1108, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 448, 640
        bmpbutton #mainWnd.cell1109, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 512, 640
        bmpbutton #mainWnd.cell1110, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 576, 640
        bmpbutton #mainWnd.cell1111, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 640, 640
        bmpbutton #mainWnd.cell1112, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 704, 640
        bmpbutton #mainWnd.cell1113, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 768, 640
        bmpbutton #mainWnd.cell1114, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 832, 640
        bmpbutton #mainWnd.cell1115, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 896, 640
        bmpbutton #mainWnd.cell1116, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 960, 640
        bmpbutton #mainWnd.cell1117, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1024, 640
        bmpbutton #mainWnd.cell1118, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1088, 640
        bmpbutton #mainWnd.cell1119, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1152, 640
        bmpbutton #mainWnd.cell1120, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1216, 640
        bmpbutton #mainWnd.cell1121, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1280, 640
        bmpbutton #mainWnd.cell1122, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1344, 640
        bmpbutton #mainWnd.cell1123, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1408, 640
        bmpbutton #mainWnd.cell1124, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1472, 640
        bmpbutton #mainWnd.cell1125, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1536, 640
        bmpbutton #mainWnd.cell1201, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 0, 704
        bmpbutton #mainWnd.cell1202, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 64, 704
        bmpbutton #mainWnd.cell1203, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 128, 704
        bmpbutton #mainWnd.cell1204, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 192, 704
        bmpbutton #mainWnd.cell1205, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 256, 704
        bmpbutton #mainWnd.cell1206, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 320, 704
        bmpbutton #mainWnd.cell1207, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 384, 704
        bmpbutton #mainWnd.cell1208, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 448, 704
        bmpbutton #mainWnd.cell1209, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 512, 704
        bmpbutton #mainWnd.cell1210, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 576, 704
        bmpbutton #mainWnd.cell1211, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 640, 704
        bmpbutton #mainWnd.cell1212, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 704, 704
        bmpbutton #mainWnd.cell1213, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 768, 704
        bmpbutton #mainWnd.cell1214, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 832, 704
        bmpbutton #mainWnd.cell1215, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 896, 704
        bmpbutton #mainWnd.cell1216, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 960, 704
        bmpbutton #mainWnd.cell1217, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1024, 704
        bmpbutton #mainWnd.cell1218, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1088, 704
        bmpbutton #mainWnd.cell1219, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1152, 704
        bmpbutton #mainWnd.cell1220, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1216, 704
        bmpbutton #mainWnd.cell1221, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1280, 704
        bmpbutton #mainWnd.cell1222, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1344, 704
        bmpbutton #mainWnd.cell1223, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1408, 704
        bmpbutton #mainWnd.cell1224, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1472, 704
        bmpbutton #mainWnd.cell1225, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1536, 704
        bmpbutton #mainWnd.cell1301, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 0, 768
        bmpbutton #mainWnd.cell1302, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 64, 768
        bmpbutton #mainWnd.cell1303, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 128, 768
        bmpbutton #mainWnd.cell1304, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 192, 768
        bmpbutton #mainWnd.cell1305, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 256, 768
        bmpbutton #mainWnd.cell1306, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 320, 768
        bmpbutton #mainWnd.cell1307, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 384, 768
        bmpbutton #mainWnd.cell1308, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 448, 768
        bmpbutton #mainWnd.cell1309, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 512, 768
        bmpbutton #mainWnd.cell1310, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 576, 768
        bmpbutton #mainWnd.cell1311, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 640, 768
        bmpbutton #mainWnd.cell1312, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 704, 768
        bmpbutton #mainWnd.cell1313, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 768, 768
        bmpbutton #mainWnd.cell1314, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 832, 768
        bmpbutton #mainWnd.cell1315, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 896, 768
        bmpbutton #mainWnd.cell1316, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 960, 768
        bmpbutton #mainWnd.cell1317, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1024, 768
        bmpbutton #mainWnd.cell1318, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1088, 768
        bmpbutton #mainWnd.cell1319, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1152, 768
        bmpbutton #mainWnd.cell1320, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1216, 768
        bmpbutton #mainWnd.cell1321, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1280, 768
        bmpbutton #mainWnd.cell1322, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1344, 768
        bmpbutton #mainWnd.cell1323, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1408, 768
        bmpbutton #mainWnd.cell1324, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1472, 768
        bmpbutton #mainWnd.cell1325, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1536, 768
        bmpbutton #mainWnd.cell1401, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 0, 832
        bmpbutton #mainWnd.cell1402, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 64, 832
        bmpbutton #mainWnd.cell1403, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 128, 832
        bmpbutton #mainWnd.cell1404, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 192, 832
        bmpbutton #mainWnd.cell1405, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 256, 832
        bmpbutton #mainWnd.cell1406, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 320, 832
        bmpbutton #mainWnd.cell1407, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 384, 832
        bmpbutton #mainWnd.cell1408, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 448, 832
        bmpbutton #mainWnd.cell1409, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 512, 832
        bmpbutton #mainWnd.cell1410, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 576, 832
        bmpbutton #mainWnd.cell1411, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 640, 832
        bmpbutton #mainWnd.cell1412, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 704, 832
        bmpbutton #mainWnd.cell1413, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 768, 832
        bmpbutton #mainWnd.cell1414, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 832, 832
        bmpbutton #mainWnd.cell1415, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 896, 832
        bmpbutton #mainWnd.cell1416, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 960, 832
        bmpbutton #mainWnd.cell1417, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1024, 832
        bmpbutton #mainWnd.cell1418, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1088, 832
        bmpbutton #mainWnd.cell1419, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1152, 832
        bmpbutton #mainWnd.cell1420, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1216, 832
        bmpbutton #mainWnd.cell1421, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1280, 832
        bmpbutton #mainWnd.cell1422, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1344, 832
        bmpbutton #mainWnd.cell1423, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1408, 832
        bmpbutton #mainWnd.cell1424, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1472, 832
        bmpbutton #mainWnd.cell1425, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1536, 832
        bmpbutton #mainWnd.cell1501, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 0, 896
        bmpbutton #mainWnd.cell1502, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 64, 896
        bmpbutton #mainWnd.cell1503, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 128, 896
        bmpbutton #mainWnd.cell1504, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 192, 896
        bmpbutton #mainWnd.cell1505, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 256, 896
        bmpbutton #mainWnd.cell1506, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 320, 896
        bmpbutton #mainWnd.cell1507, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 384, 896
        bmpbutton #mainWnd.cell1508, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 448, 896
        bmpbutton #mainWnd.cell1509, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 512, 896
        bmpbutton #mainWnd.cell1510, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 576, 896
        bmpbutton #mainWnd.cell1511, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 640, 896
        bmpbutton #mainWnd.cell1512, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 704, 896
        bmpbutton #mainWnd.cell1513, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 768, 896
        bmpbutton #mainWnd.cell1514, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 832, 896
        bmpbutton #mainWnd.cell1515, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 896, 896
        bmpbutton #mainWnd.cell1516, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 960, 896
        bmpbutton #mainWnd.cell1517, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1024, 896
        bmpbutton #mainWnd.cell1518, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1088, 896
        bmpbutton #mainWnd.cell1519, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1152, 896
        bmpbutton #mainWnd.cell1520, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1216, 896
        bmpbutton #mainWnd.cell1521, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1280, 896
        bmpbutton #mainWnd.cell1522, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1344, 896
        bmpbutton #mainWnd.cell1523, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1408, 896
        bmpbutton #mainWnd.cell1524, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1472, 896
        bmpbutton #mainWnd.cell1525, "bmp\"; gUnvisitedBmp$; ".bmp", handleGridClick, UL, 1536, 896
    end sub

    ' ------------------------------------------------------------------------------------------------------------
    '
    ' Library code from external sources
    '
    ' ------------------------------------------------------------------------------------------------------------

    ' From QB64 and found for me by + thanks
    Function strReplace$(s$, replace$, new$) 'case sensitive QB64 2020-07-28 version
        If Len(s$) = 0 Or Len(replace$) = 0 Then
            strReplace$ = s$: Exit Function
        Else
            LR = Len(replace$): lNew = Len(new$)
        End If
        sCopy$ = s$ ' otherwise s$ would get changed in qb64
        p = InStr(sCopy$, replace$)
        While p
            sCopy$ = Mid$(sCopy$, 1, p - 1) + new$ + Mid$(sCopy$, p + LR)
            p = InStr(sCopy$, replace$, p + lNew) ' InStr is differnt in JB
        Wend
        strReplace$ = sCopy$
    End Function

    ' From "jb20help" html
    function fileExists(path$, filename$)
        files path$, filename$, info$()
        fileExists = val(info$(0, 0)) 'non zero is true
    end function

    ' Somewhat modified by me to add keyboard shortcuts
    function PopupMenu$(options$, width, bgColor$, textColor$, selBackColor$, selTextColor$, ULx, ULy, title$)
        'arguments:
        'options$ - comma-separated list of menu options
        'width - window-width, default = 100
        'bgColor$ - background color of the dialog
        'textColor$ - color of inactive text
        'selBackColor$ - backcolor of active, selected text
        'selTextColor$ - color of active, selected text
        'NOTE: colors are either a string of rgb values, one of the windows colours or
            'empty string (use default colour scheme)
        while word$(options$, count+1, ",") <> ""
            count = count+1
        wend
        height = count*20+38
        width = int(width) : if width < 100 then width = 100
        if bgColor$ = "" then bgColor$ = "white"
        if textColor$ = "" then textColor$ = "black"
        if selBackColor$ = "" then selBackColor$ = "darkblue"
        if selTextColor$ = "" then selTextColor$ = "white"
        WindowHeight = height
        WindowWidth = width
        UpperLeftX = ULx
        UpperLeftY = ULy
        graphicbox #popup.graph, 0, 0, width, height
        open title$ for dialog_modal as #popup
        #popup, "trapclose [popupDlgCancel]"
        #popup, "font ms_sans_serif 16 9"
        #popup.graph, "down; fill "; bgColor$
        #popup.graph, "color "; textColor$; "; backcolor "; bgColor$
        for i = 1 to count
            #popup.graph, "place 4 "; i*20 - 2
            #popup.graph, "\"; word$(options$, i, ",")
        next i
        this = 1
        selection = 1
        gosub [select]
        ignore2ndKeyReturnValue = 0
        #popup.graph, "flush"
        #popup.graph, "when mouseMove [popupDlgMove]"
        #popup.graph, "when leftButtonDown [popupDlgSelect]"
        #popup.graph, "when characterInput [popupDefault]"
        wait

    [popupDlgMove]
        this = (MouseY-3)/20 : if this >= 0 then this = this + 1
        this = int(this)
        if this <> selection then
            gosub [unselect]
            if this > 0 and this <= count then
                gosub [select]
            end if
            selection = this
        end if
        wait

    [unselect]
        #popup.graph, "backcolor "; bgColor$; "; color "; bgColor$
        #popup.graph, "place 2 "; selection*20 - 16; "; boxfilled "; width-12; " "; selection*20+2
        #popup.graph, "color "; textColor$
        #popup.graph, "place 4 "; selection*20 - 2
        #popup.graph, "\"; word$(options$, selection, ",")
        return

    [select]
        #popup.graph, "backcolor "; selBackColor$; "; color "; selBackColor$
        #popup.graph, "place 2 "; this*20 - 16; "; boxfilled "; width-12; " "; this*20+2
        #popup.graph, "color "; selTextColor$
        #popup.graph, "place 4 "; this*20 - 2
        #popup.graph, "\"; word$(options$, this, ",")
        return

    [popupDlgSelect]
        this = (MouseY-3)/20 : if this >= 0 then this = this + 1
        this = int(this)
        if this > count or this < 1 then wait
        PopupMenu$ = word$(options$, this, ",")
        goto [popupDlgCancel]

    [popupDefault]
        if not(ignore2ndKeyReturnValue) then
            ignore2ndKeyReturnValue = 1
            select case asc(right$(Inkey$, 1))
                case _VK_CONTROL
                    PopupMenu$ = word$(options$, selection, ",")
                    goto [popupDlgCancel]
                case _VK_SHIFT
                    gosub [unselect]
                    this = selection + 1
                    if this > count then this = 1
                    gosub [select]
                    selection = this
            end select
        else
            ignore2ndKeyReturnValue = 0
        end if
        wait

    [popupDlgCancel]
        close #popup
    end function


