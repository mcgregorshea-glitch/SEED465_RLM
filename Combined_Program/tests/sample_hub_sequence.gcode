; Sample G-Code with embedded Vivigo Hub commands
G28 ; Home everything
G1 X110 Y110 Z50 F3000 ; Move to center
; HUB_CMD param=WPT Start value=true
; WAIT seconds=2
G1 X120 Y120 F1000 ; Move to offset
; HUB_CMD param=WPT Stop value=true
G1 X110 Y110 F3000 ; Return to center
