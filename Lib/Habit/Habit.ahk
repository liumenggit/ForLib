Habit(HabitFUN)
{
	HKL:=DllCall("LoadKeyboardLayout", Str, HabitFUN.Ĭ�����뷨.���뷨������, UInt, 1)
	ControlGetFocus,ctl,A
	SendMessage,0x50,0,HKL,%ctl%,A
	ime:=HabitFUN.Ĭ�����뷨.����1Ӣ��0
	ptrSize := !A_PtrSize ? 4 : A_PtrSize
	VarSetCapacity(stGTI, cbSize:=4+4+(PtrSize*6)+16, 0)
	NumPut(cbSize, stGTI,  0, "UInt")   ;	DWORD   cbSize;
	DllCall("GetGUIThreadInfo", Uint,0, Uint,&stGTI)
	hwnd :=  NumGet(stGTI,8+PtrSize,"UInt")
	DllCall("SendMessage", "UInt", DllCall("imm32\ImmGetDefaultIMEWnd", "UInt", HWND), "UInt", 0x0283,  "Int", 0x006,  "Int",ime)
}
