    nomainwin

    WindowWidth = 1800
    WindowHeight = 1100
    UpperLeftX = 0
    UpperLeftY = 0
    open "Drawing" for graphics as #handle
    #handle "trapclose [quit]"

    loadbmp "splash", "counter scans\splash.bmp"
    #handle "drawbmp splash 0 0"
    for row = 1 to 15
        yUL$ = str$((row - 1) * 64)
        row$ = str$(row)
        if row < 10 then row$ = "0"; str$(row)
        for col = 1 to 25
            xUL$ = str$((col - 1) * 64)
            col$ = str$(col)
            if col < 10 then col$ = "0"; str$(col)
            #handle "getbmp test "; xUL$; " "; yUL$; " 64 64"
            bmpsave "test", "bmp\splash\"; row$; col$; ".bmp"
            UNLOADBMP "test"
        next col
    next row
    wait

[quit]
    unloadBmp "splash"
    close #handle
    end
