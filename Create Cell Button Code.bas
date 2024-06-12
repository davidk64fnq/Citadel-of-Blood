    gridSize = 64

    print "    sub createButtons"
    for row = 1 to 15
        gridPosY = (row - 1) * gridSize
        for col = 1 to 25
            ' pad single digit grid numbers
            row$ = str$(row)
            if row < 10 then row$ = "0"; str$(row)
            col$ = str$(col)
            if col < 10 then col$ = "0"; str$(col)
            buttonHandle$ = "#mainWnd.cell" + row$ + col$

            gridPosX = (col - 1) * gridSize
            print "        bmpbutton " + buttonHandle$ + ", " + chr$(34) + "bmp\nnnnn10n64.bmp" + chr$(34);
            print ", handleGridClick, UL, " + str$(gridPosX) + ", " + str$(gridPosY)
        next col
    next row
    print "    end sub"
