    global gFeatureCodes$, gFeatures$, gNoFeatures,_
           gDoorWidth, gDoorWidth$, gTileSize, gTileWidth$, gFeatureUL$,_
           gFsize$, gCsize$, gDsize$, gTsize$,_
           gExitTypes$, gCorridor$, gDoor$, gWall$, gRoom$, gUnknown$, gNormal$, gPartyLoc$

    gFeatures$ = "Altar,Artwork,Fountain,Furniture,Mirror,Staircase,Statue,Trap,Empty,Hellgate,Gateway"
    gFeatureCodes$ = "01 02 03 04 05 06 07 08 09 10 11"
    gNoFeatures = 11
    gExitTypes$ = "c d w"
    gTsize$ = "32 48 64"    ' Tile bitmap dimensions
    gFsize$ = "18 27 36"    ' Feature bitmap dimensions
    gCsize$ = "20 30 40"    ' Corridor (and room) bitmap dimensions
    gDsize$ = "6 9 12"      ' Door bitmap dimensions
    gRoom$ = "r"        ' Segment tile type
    gCorridor$ = "c"    ' Can be segment tile type or segment exit type
    gDoor$ = "d"        ' Segment exit type
    gWall$ = "w"        ' Segment (non) exit type
    gUnknown$ = "u"     ' Used with Hellgate and Unvisited segments for tile type and exit type
    gNormal$ = "n"      ' Bitmap name 10th char denotes a bitmap showing party NOT at that location
    gPartyLoc$ = "p"    ' Bitmap name 10th char denotes a bitmap showing party at that location

    open "Drawing" for graphics as #handle
    #handle "trapclose [quit]"

    call loadBitmaps

    ' Three versions 32x32, 48x48 and 64x64
    for factor = 0.5 to 1.0 step 0.25
        ' For each exit direction three exit options, door, corridor or none
        for eastIndex = 1 to 3
            for southIndex = 1 to 3
                for westIndex = 1 to 3
                    for northIndex = 1 to 3
                        exits$ = word$(gExitTypes$, eastIndex);word$(gExitTypes$, southIndex);_
                                word$(gExitTypes$, westIndex);word$(gExitTypes$, northIndex)
                        ' Exclude the tile with no exits
                        if exits$ <> "wwww" then
                            ' Rooms have no corridor exits
                            if not(instr(exits$, gCorridor$)) then
                                ' Version for each feature for room tiles (exclude Hellgate and Gateway)
                                for feature = 1 to gNoFeatures - 2
                                    featureCode$ = word$(gFeatureCodes$, feature)
                                    ' Two version, one for party not at that tile, the other they are
                                    call createTile factor, gRoom$; exits$; featureCode$; gNormal$
                                    call createTile factor, gRoom$; exits$; featureCode$; gPartyLoc$
                                next feature
                            else
                                ' Corridors don't have features - just empty space in tile centre
                                emptyFeature$ = getFeatureNo$("Empty")
                                call createTile factor, gCorridor$; exits$; emptyFeature$; gNormal$
                                call createTile factor, gCorridor$; exits$; emptyFeature$; gPartyLoc$

                                ' Gateway tiles
                                if instr("cwww wcww wwcw wwwc", exits$) then
                                    call createTile factor, gCorridor$; exits$; getFeatureNo$("Gateway"); gNormal$
                                    call createTile factor, gCorridor$; exits$; getFeatureNo$("Gateway"); gPartyLoc$
                                end if
                            end if
                        end if
                    next northIndex
                next westIndex
            next southIndex
        next eastIndex
    next factor

    ' Hellgate tiles and Background tiles
    for factor = 0.5 to 1.0 step 0.25
        call setGlobals factor
        call drawUnvisited factor
        call saveTile factor, "uuuuu"; getFeatureNo$("Empty"); gNormal$

        call drawBackground factor
        call drawFeature factor, "Hellgate"
        call saveTile factor, "uuuuu"; getFeatureNo$("Hellgate"); gNormal$

        call drawBackground factor
        call drawFeature factor, "Hellgate"
        call drawPartyIndicator factor
        call saveTile factor, "uuuuu"; getFeatureNo$("Hellgate"); gPartyLoc$
    next factor

    wait

[quit]
    call unloadBitmaps
    close #handle
    end

    sub loadBitmaps
        ' Load floor bitmaps
        for bSize = 1 to 3
            bmpName$ = "Background "
            bmpSize$ = word$(gTsize$, bSize); "x"; word$(gTsize$, bSize)
            call loadBitmap bmpName$, bmpSize$

            bmpName$ = "Background Room "
            bmpSize$ = word$(gCsize$, bSize); "x"; word$(gCsize$, bSize)
            call loadBitmap bmpName$, bmpSize$

            bmpName$ = "Background Door "
            bmpSize$ = word$(gDsize$, bSize); "x"; word$(gDsize$, bSize)
            call loadBitmap bmpName$, bmpSize$

            bmpName$ = "Background Corridor NS "
            bmpSize$ = word$(gDsize$, bSize); "x"; word$(gCsize$, bSize)
            call loadBitmap bmpName$, bmpSize$

            bmpName$ = "Background Corridor EW "
            bmpSize$ = word$(gCsize$, bSize); "x"; word$(gDsize$, bSize)
            call loadBitmap bmpName$, bmpSize$
        next bSize

        ' Load feature bitmaps
        for index = 1 to gNoFeatures
            bmpFeature$ = word$(gFeatures$, index, ","); " "
            for fSize = 1 to 3
                bmpSize$ = word$(gFsize$, fSize); "x"; word$(gFsize$, fSize)
                call loadBitmap bmpFeature$, bmpSize$
            next fSize
        next index
    end sub

    sub loadBitmap bmpName$, bmpSize$
        fileName$ = bmpName$; bmpSize$
        bmpName$ = strReplace$(fileName$, " ", "")
        loadbmp bmpName$, "counter scans\";fileName$;".bmp"
    end sub

    sub createTile factor, bitmapName$
        call setGlobals factor
        call drawBlank factor
        call drawRoom factor
        for direction = 1 to 4
            exitType$ = mid$(bitmapName$, direction + 1, 1)
            select case str$(direction);exitType$
                case "1c"
                    call drawCorridorEast factor
                case "1d"
                    call drawDoorEast factor
                case "2c"
                    call drawCorridorSouth factor
                case "2d"
                    call drawDoorSouth factor
                case "3c"
                    call drawCorridorWest factor
                case "3d"
                    call drawDoorWest factor
                case "4c"
                    call drawCorridorNorth factor
                case "4d"
                    call drawDoorNorth factor
            end select
        next direction
        featureName$ = word$(gFeatures$, val(mid$(bitmapName$, 6, 2)), ",")
        call drawFeature factor, featureName$
        if right$(bitmapName$, 1) = gPartyLoc$  then
            call drawPartyIndicator factor
        end if

        call saveTile factor, bitmapName$
    end sub

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

    sub setGlobals factor
        gDoorWidth = 12 * factor
        gTileSize = 64 * factor
        gTileWidth$ = str$(gTileSize)
        gDoorWidth$ = str$(gDoorWidth)
        gFeatureUL$ = str$(gTileSize/2 - (36 * factor)/2)
    end sub

    sub drawBlank factor
        #handle "cls"

        ' light trim outside
        #handle "size 1; color 58 159 255"
        #handle "down"
        #handle "line 0 0 "; gTileWidth$; " 0"
        #handle "line 0 0 0 "; gTileWidth$
        #handle "up"

        ' light trim inside
        #handle "size 1; color 47 129 207"
        #handle "down"
        #handle "line 1 1 "; str$(gTileSize - 1); " 1"
        #handle "line 1 1 1 "; str$(gTileSize - 1)
        #handle "up"

        ' dark trim inside
        #handle "size 1; color 28 76 122"
        #handle "down"
        #handle "line 1 "; str$(gTileSize - 2); " "; str$(gTileSize - 1); " "; str$(gTileSize - 2)
        #handle "line "; str$(gTileSize - 2); " 1 "; str$(gTileSize - 2); " "; str$(gTileSize - 1)
        #handle "up"

        ' dark trim outside
        #handle "goto 2 2"
        #handle "size 1; color 22 62 99"
        #handle "down"
        #handle "line 0 "; str$(gTileSize - 1); " "; gTileWidth$; " "; str$(gTileSize - 1)
        #handle "line "; str$(gTileSize - 1); " 0 "; str$(gTileSize - 1); " "; gTileWidth$
        #handle "up"

        ' background
        #handle "goto 2 2"
        #handle "color 33 91 146; backcolor 33 91 146"
        #handle "down"
        #handle "boxfilled "; str$(gTileSize - 2); " "; str$(gTileSize - 2)
        #handle "up"
    end sub

    sub drawRoom factor
        #handle "drawbmp "; getBmp$(factor, "BackgroundRoom"); " "; gDoorWidth$; " "; gDoorWidth$
    end sub

    sub drawCorridorEast factor
        call drawRoom factor
        #handle "drawbmp "; getBmp$(factor, "BackgroundCorridorNS"); " "; str$(gTileSize - gDoorWidth); " "; gDoorWidth$
    end sub

    sub drawDoorEast factor
        #handle "drawbmp "; getBmp$(factor, "BackgroundDoor"); " "; str$(gTileSize - gDoorWidth); " "; str$(gTileSize/2 - gDoorWidth/2)
    end sub

    sub drawCorridorSouth factor
        call drawRoom factor
        #handle "drawbmp "; getBmp$(factor, "BackgroundCorridorEW"); " "; gDoorWidth$; " "; str$(gTileSize - gDoorWidth)
    end sub

    sub drawDoorSouth factor
        #handle "drawbmp "; getBmp$(factor, "BackgroundDoor"); " "; str$(gTileSize/2 - gDoorWidth/2); " "; str$(gTileSize - gDoorWidth)
    end sub

    sub drawCorridorWest factor
        call drawRoom factor
        #handle "drawbmp "; getBmp$(factor, "BackgroundCorridorNS"); " 0 "; gDoorWidth$
    end sub

    sub drawDoorWest factor
        #handle "drawbmp "; getBmp$(factor, "BackgroundDoor"); " 0 "; str$(gTileSize/2 - gDoorWidth/2)
    end sub

    sub drawCorridorNorth factor
        call drawRoom factor
        #handle "drawbmp "; getBmp$(factor, "BackgroundCorridorEW"); " "; gDoorWidth$; " 0"
    end sub

    sub drawDoorNorth factor
        #handle "drawbmp "; getBmp$(factor, "BackgroundDoor"); " "; str$(gTileSize/2 - gDoorWidth/2); " 0"
    end sub

    sub drawFeature factor, featureName$
        #handle "drawbmp "; getBmp$(factor, featureName$); " "; gFeatureUL$; " "; gFeatureUL$
    end sub

    function getBmp$(factor, type$)
        bIndex = (factor - 0.25) / 0.25
        select case type$
            case "Background"
                bmpSize$ = word$(gTsize$, bIndex); "x"; word$(gTsize$, bIndex)
            case "BackgroundDoor"
                bmpSize$ = word$(gDsize$, bIndex); "x"; word$(gDsize$, bIndex)
            case "BackgroundRoom"
                bmpSize$ = word$(gCsize$, bIndex); "x"; word$(gCsize$, bIndex)
            case "BackgroundCorridorEW"
                bmpSize$ = word$(gCsize$, bIndex); "x"; word$(gDsize$, bIndex)
            case "BackgroundCorridorNS"
                bmpSize$ = word$(gDsize$, bIndex); "x"; word$(gCsize$, bIndex)
            case else
                bmpSize$ = word$(gFsize$, bIndex); "x"; word$(gFsize$, bIndex)
        end select
        getBmp$ = type$; bmpSize$
    end function

    sub drawPartyIndicator factor
        #handle "goto 4 4"
        #handle "size 1; color yellow"
        #handle "down"
        #handle "box "; str$(gTileSize - 4); " "; str$(gTileSize - 4)
        #handle "up"
    end sub

    sub saveTile factor, name$
        #handle "getbmp test 0 0 "; gTileWidth$; " "; gTileWidth$
        bmpsave "test", "bmp\"; name$; gTileWidth$; ".bmp"
        unloadbmp "test"
    end sub

    sub drawBackground factor
        #handle "drawbmp "; getBmp$(factor, "Background"); " 0 0"
    end sub

    sub drawUnvisited factor
        #handle "goto 0 0"
        #handle "color 153 217 234; backcolor 153 217 234"
        #handle "down"
        #handle "boxfilled "; gTileWidth$; " "; gTileWidth$
        #handle "up"
    end sub

    sub unloadBitmaps
        ' unload floor bitmaps
        for bSize = 1 to 3
            bmpName$ = "Background "
            bmpSize$ = word$(gTsize$, bSize); "x"; word$(gTsize$, bSize)
            call unloadBitmap bmpName$, bmpSize$

            bmpName$ = "Background Room "
            bmpSize$ = word$(gCsize$, bSize); "x"; word$(gCsize$, bSize)
            call unloadBitmap bmpName$, bmpSize$

            bmpName$ = "Background Door "
            bmpSize$ = word$(gDsize$, bSize); "x"; word$(gDsize$, bSize)
            call unloadBitmap bmpName$, bmpSize$

            bmpName$ = "Background Corridor NS "
            bmpSize$ = word$(gDsize$, bSize); "x"; word$(gCsize$, bSize)
            call unloadBitmap bmpName$, bmpSize$

            bmpName$ = "Background Corridor EW "
            bmpSize$ = word$(gCsize$, bSize); "x"; word$(gDsize$, bSize)
            call unloadBitmap bmpName$, bmpSize$
        next bSize

        ' unload feature bitmaps
        for index = 1 to gNoFeatures
            bmpFeature$ = word$(gFeatures$, index, ","); " "
            for fSize = 1 to 3
                bmpSize$ = word$(gFsize$, fSize); "x"; word$(gFsize$, fSize)
                call unloadBitmap bmpFeature$, bmpSize$
            next fSize
        next index
    end sub

    sub unloadBitmap bmpName$, bmpSize$
        fileName$ = bmpName$; bmpSize$
        bmpName$ = strReplace$(fileName$, " ", "")
        unloadbmp bmpName$
    end sub

    ' Library code from external sources

    ' From QB64 and found for me by + thanks
    Function strReplace$(s$, replace$, new$) 'case sensitive  QB64 2020-07-28 version
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

