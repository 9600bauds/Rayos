global vModifArticulo_id := "Modificación : ahk_exe LUPA.exe"
global vNuevoArticulo_id := "Nuevo"

global vModifArticulo_codigo := "Edit1"
global vModifArticulo_codigoBarras := "Edit2"
global vModifArticulo_descripcion := "Edit3"
global vModifArticulo_descripcionReducida := "Edit4"
global vModifArticulo_puntoPedido := "Edit5"
global vModifArticulo_empaque := "Edit6"
global vModifArticulo_unidad := "ComboBox1"
global vModifArticulo_moneda := "ComboBox2"
global vModifArticulo_precioCosto := "Edit7"
global vModifArticulo_margen1 := "Edit8"
global vModifArticulo_margen2 := "Edit11"
global vModifArticulo_margen3:= "Edit14"
global vModifArticulo_iva := "ComboBox3"
global vModifArticulo_rubro := "ComboBox4"
global vModifArticulo_nota := "Edit17"
global vModifArticulo_ok := "Button18"
global vModifArticulo_cancela := "Button19"

vModifArticulo_Cerrar(){
	if(!WinExist(vModifArticulo_id)) return
		
	ControlFocus, %vModifArticulo_cancela%, %vModifArticulo_id%,,,, NA
	ControlClick, %vModifArticulo_cancela%, %vModifArticulo_id%,,,, NA
	ControlSend, %vModifArticulo_cancela%, {Enter}, %vModifArticulo_id%
	WinKill, %vModifArticulo_id%
	WaitControlNotExist(vModifArticulo_cancela, vModifArticulo_id) 
}

vModifArticulo_Abrir(){
	if(WinExist(vModifArticulo_id)) return

	Sleep, 200
	
	ControlFocus, %vReporteArticulos_modificar%, %vReporteArticulos_id%
	Sleep, 200
	ControlClick, %vReporteArticulos_modificar%, %vReporteArticulos_id%
	;ControlSend, %vReporteArticulos_modificar%, ^M, %vReporteArticulos_id%
	Loop{
		if(A_Index = 350){
			MsgBox, GetAlias - Could not get (551) vModifArticulo_id.
			return
		}
		;ControlClick, %vReporteArticulos_modificar%, %vReporteArticulos_id%,,,, NA
		;ControlSend, %vReporteArticulos_modificar%, {Space}, %vReporteArticulos_id%
		WinWait, %vModifArticulo_id%, , 0.5
		if not ErrorLevel {
			Break
		}
	}
}