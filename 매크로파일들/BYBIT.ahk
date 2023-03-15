Gui, Add, Text, x30 y5 w200 h20, BTBIT AUTOMATIC
Gui, Add, Text, x30 y35 w110 h20, F 1 0 : S T A R T
Gui, Add, Text, x30 y55 w110 h20, F 1 2 : E x i t
Gui, Add, Button, x20 y80 w110 h20, Start

Gui, show

URL = https://www.bybit.com/en-US/web3
;=======2022.11.11.Ver.01=======

F10::
{
send, !d
send, https://www.bybit.com/en-US/web3
Sleep 1000
send,{Enter} ; 바이빗 이동
send,{F11}
Sleep 2500
MouseClick, left, 1746 , 24 , 1, 0 ; 로그인 클릭
Sleep 1500
MouseClick, left, 865 , 718 , 1, 0
Sleep 2500
send, {Tab}
send, {Tab}
send, {Enter}
}

