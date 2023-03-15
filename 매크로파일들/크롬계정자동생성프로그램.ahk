Gui, Add, Text, x30 y5 w200 h20, Chrome Account Auto Create Ready
Gui, Add, Text, x30 y35 w110 h20, F 1 2 : E x i t
Gui, Add, Text, x30 y55 w110 h20, Made by SBU...
Gui, Add, Button, x20 y80 w110 h20, Start

;Gui, Add, Edit, x22 y59 w100 h20 vEdit,
;Gui, Show, w146 h145,GUI

;Gui, Add, Button, x20 y110 w110 h20, End
Gui, show
Start := false ;추후에 입력값 받는거 구현하기

;===========컨트롤러===========

n = 1001 ;시작 숫자 설정
j = 1000 ;반복 횟수 설정
forcount = 1
loopcount = n
JUMP := 9
;JUMP : 사용자 선택 시 왼쪽부터
;n = n/9지점 클릭


;=======2022.11.11.Ver.01=======


ButtonStart:
Loop,%j%
{

{
	if(Start == false)
	{
			CoordMode,Mouse,Window
			Sleep 500
			Send, #d ; 화면정리
			MouseClick, left, 0 , 0 , 1, 0 ;영점 이동
			MouseClick, left, 40 , 40 , 1, 0 ; 크롬 바로가기 클릭
			MouseClick, Left
			Send,{Enter} ; 크롬 들어가기

			if(loopcount > 500)
			{
				Sleep 5000
			}
			if(loopcount > 700)
			{
				Sleep 5500
			}
			else
			{
				Sleep 4500
			}

			Send,{F11} ; 전체화면
			Sleep 500
			;CoordMode,Mouse,Window
			MouseClick, left, 901, 77 , 1 , 0 ;스크롤 위해 중앙부분 클릭
	}

	if(Start == false)
	{
			Sleep 500
			MouseClickDrag, left, 1778, 203 , 1778, 953 ;스크롤 다운
	}

	if(Start == false)
	{
		sleep, 500
		if(JUMP == 1)
		{
			MouseClick, left, 248, 871 , 1 , 0 ;2/9지점 클릭 ;1/9지점 클릭
			sleep, 500
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Enter} ; 계정없이 계속 들어감
			JUMP := 2
			Goto, A
		}

		if(JUMP == 2)
		{
			MouseClick, left, 425, 871 , 1 , 0 ;2/9지점 클릭
			sleep, 500
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Enter} ; 계정없이 계속 들어감
			JUMP := 3
			Goto, B
		}

		if(JUMP == 3)
		{
			MouseClick, left, 608, 871 , 1 , 0 ;3/9지점 클릭
			sleep, 500
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Enter} ; 계정없이 계속 들어감
			JUMP := 4
			Goto, C
		}

		if(JUMP == 4)
		{
			MouseClick, left, 710, 871 , 1 , 0 ;4/9지점 클릭
			sleep, 500
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Enter} ; 계정없이 계속 들어감
			JUMP := 5
			Goto, D
		}

		if(JUMP == 5)
		{
			MouseClick, left, 966, 871 , 1 , 0 ;5/9지점 클릭
			sleep, 500
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Enter} ; 계정없이 계속 들어감
			JUMP := 6
			Goto, E
		}

		if(JUMP == 6)
		{
			MouseClick, left, 1151, 871 , 1 , 0 ;6/9지점 클릭
			sleep, 500
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Enter} ; 계정없이 계속 들어감
			JUMP := 7
			Goto, F
		}

		if(JUMP == 7)
		{
			MouseClick, left, 1311, 871 , 1 , 0 ;7/9지점 클릭
			sleep, 500
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Enter} ; 계정없이 계속 들어감
			JUMP := 8
			Goto, G
		}

		if(JUMP == 8)
		{
			MouseClick, left, 1500, 871 , 1 , 0 ;8/9지점 클릭
			sleep, 500
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Enter} ; 계정없이 계속 들어감
			JUMP := 9
			Goto, H
		}

		if(JUMP == 9)
		{
			MouseClick, left, 1685, 871 , 1 , 0 ;9/9지점 클릭
			sleep, 500
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Tab} ; 계정없이 계속
			Sleep 50
			Send,{Enter} ; 계정없이 계속 들어감
			JUMP := 1
			Goto, I
		}
	}

	A:
	B:
	C:
	D:
	E:
	F:
	G:
	H:
	I:


	if(Start == false) ;나중에 for문으로 크게 묶자
	{
		if(n <= 9) {
			Send,000
		}else if(n <= 99){
			Send,00
		}else if(n <= 999){
			Send,0
		}else{
			Sleep 1
		}
		Send,%n%
		sleep, 500
		Send,{Tab} ; 완료
		Sleep 300
		Send,{Tab} ; 완료
		Sleep 300
		Send,{Tab} ; 완료
		Sleep 300
		Send,{Enter} ; 완료
		Sleep 1500
		Send !{ f4 } ; 크롬종료
		Sleep 500
		Send, #d ; 화면정리
		n++ ;n에 1 추가하여 다시 반복

	}

	if(Start == false) ;나중에 for문으로 크게 묶자
	{
		CoordMode,Mouse,Window
		Sleep 500
		if(forcount == 50)
		{
			Sleep 500
			MouseClick, left, 1919 , 0 , 1, 0 ; 영점클릭
			MouseClickDrag, left, 1888, 19 , 75, 1000 ;스크롤 다운
			Sleep 200
			Send, ^x
			Sleep 200
			MouseClick, left, 40 , 140 , 1, 0 ; 크롬 바로가기 클릭
			MouseClick, Left
			Send,{Enter} ; 크롬 들어가기
			Sleep 500
			Send, ^v
			Sleep 2000
			Send !{ f4 }
			forcount = 1
		}
		MouseClick, left, -1000 , -1000 , 1, 0 ;영점 이동
		forcount++
		loopcount++
	}
}

}

F12::
ExitApp