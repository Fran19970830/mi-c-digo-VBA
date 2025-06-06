Private Sub Cerrar_Click()
    Unload Me
End Sub

Private Sub Describa_Accesibilidad_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Edificio_Change()
    Edificio_Inventario.Text = Edificio.Text
End Sub

Private Sub Edificio_Inventario_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Edificio_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Guardar_Click()
Sheets("Registro").Select
Dim aux As Double
aux = 0
           
If aux < 1 Then
        Dim contador As Integer
        Range("A1000000").End(xlUp).Offset(1, 0).Select
        ActiveCell.Value = RFI
        ActiveCell.Offset(0, 1) = Edificio_Inventario
        ActiveCell.Offset(0, 2) = Instituto
        ActiveCell.Offset(0, 3) = Edificio
        ActiveCell.Offset(0, 4) = Ubicacion
        ActiveCell.Offset(0, 5) = Actividad
        ActiveCell.Offset(0, 6) = Tipo_Propiedad
        ActiveCell.Offset(0, 7) = Sector
        ActiveCell.Offset(0, 8) = Sup_Terreno
        ActiveCell.Offset(0, 9) = Sup_Total
        ActiveCell.Offset(0, 10) = Num_Edificaciones
        ActiveCell.Offset(0, 11) = Folio
        ActiveCell.Offset(0, 12) = Num_Niveles
        ActiveCell.Offset(0, 13) = Num_Cuerpos
        ActiveCell.Offset(0, 14) = Num_Edif_Sin_Acc
        ActiveCell.Offset(0, 15) = Usuario_Cedula
        ActiveCell.Offset(0, 16) = Año_Construccion
        ActiveCell.Offset(0, 17) = Remodelacion
        ActiveCell.Offset(0, 18) = Tipo_Inmueble
        ActiveCell.Offset(0, 19) = Fecha_Registro
        
        If Aplica_SI.Value = True Then
            ActiveCell.Offset(0, 20) = 1
            ActiveCell.Offset(0, 21) = 0
        ElseIf Aplica_NO.Value = True Then
            ActiveCell.Offset(0, 20) = 0
            ActiveCell.Offset(0, 21) = 1
        Else
            MsgBox ("SELECCIONA UNA OPCIÓN QUE APLICA LA ACCESIBILIDAD")
        End If
         
        ActiveCell.Offset(0, 22) = Total_Personal
        ActiveCell.Offset(0, 23) = Poblacion_Servicio
        ActiveCell.Offset(0, 24) = Estacionamiento_Publico
        ActiveCell.Offset(0, 25) = Estacionamiento_Personal
        ActiveCell.Offset(0, 26) = Elevador
        ActiveCell.Offset(0, 27) = Escaleras
        ActiveCell.Offset(0, 28) = Salida_Emergencia
        ActiveCell.Offset(0, 29) = Personal_Discapacidad
        ActiveCell.Offset(0, 30) = Poblacion_A_Diario
        ActiveCell.Offset(0, 31) = Cajon_Discapacidad1
        ActiveCell.Offset(0, 32) = Cajon_Discapacidad2
        ActiveCell.Offset(0, 33) = Salva_Escalera
        ActiveCell.Offset(0, 34) = Escalera_Emergencia
        ActiveCell.Offset(0, 35) = Desnivel_Acceso

        'INDICADORES DE ACCESIBILIDAD
        'ACCESO AL INMUEBLE

        If OB_SI1.Value = True Then
            ActiveCell.Offset(0, 36) = 1
            ActiveCell.Offset(0, 37) = 0
        ElseIf OB_NO1.Value = True Then
            ActiveCell.Offset(0, 36) = 0
            ActiveCell.Offset(0, 37) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON EL ALUMBRADO PÚBLICO EN EL ACCESO")
        End If
        
        If OB_SI2.Value = True Then
            ActiveCell.Offset(0, 38) = 1
            ActiveCell.Offset(0, 39) = 0
        ElseIf OB_NO2.Value = True Then
            ActiveCell.Offset(0, 38) = 0
            ActiveCell.Offset(0, 39) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON EL ALUMBRADO PÚBLICO EN EL ACCESO")
        End If
        
        If OB_SI3.Value = True Then
            ActiveCell.Offset(0, 40) = 1
            ActiveCell.Offset(0, 41) = 0
        ElseIf OB_NO3.Value = True Then
            ActiveCell.Offset(0, 40) = 0
            ActiveCell.Offset(0, 41) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON LA RAMPA EN ACCESOS CON PENDIENTES DE 6%")
        End If
        
        If OB_SI4.Value = True Then
            ActiveCell.Offset(0, 42) = 1
            ActiveCell.Offset(0, 43) = 0
        ElseIf OB_NO4.Value = True Then
            ActiveCell.Offset(0, 42) = 0
            ActiveCell.Offset(0, 43) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON LA RAMPA EN ACCESOS CON PENDIENTES DE 6%")
        End If
        
        If OB_SI5.Value = True Then
            ActiveCell.Offset(0, 44) = 1
            ActiveCell.Offset(0, 45) = 0
        ElseIf OB_NO5.Value = True Then
            ActiveCell.Offset(0, 44) = 0
            ActiveCell.Offset(0, 45) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON PUERTAS DE ACCESO DE 1.20 MTS. ANCHO")
        End If
        
        If OB_SI6.Value = True Then
            ActiveCell.Offset(0, 46) = 1
            ActiveCell.Offset(0, 47) = 0
        ElseIf OB_NO6.Value = True Then
            ActiveCell.Offset(0, 46) = 0
            ActiveCell.Offset(0, 47) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON PUERTAS DE ACCESO DE 1.20 MTS. ANCHO")
        End If
        
        If OB_SI7.Value = True Then
            ActiveCell.Offset(0, 48) = 1
            ActiveCell.Offset(0, 49) = 0
        ElseIf OB_NO7.Value = True Then
            ActiveCell.Offset(0, 48) = 0
            ActiveCell.Offset(0, 49) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON LA SEÑALIZACIÓN VISUAL EN EL ACCESO")
        End If
        
        If OB_SI8.Value = True Then
            ActiveCell.Offset(0, 50) = 1
            ActiveCell.Offset(0, 51) = 0
        ElseIf OB_NO8.Value = True Then
            ActiveCell.Offset(0, 50) = 0
            ActiveCell.Offset(0, 51) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON LA SEÑALIZACIÓN VISUAL EN EL ACCESO")
        End If
        
        If OB_SI9.Value = True Then
            ActiveCell.Offset(0, 52) = 1
            ActiveCell.Offset(0, 53) = 0
        ElseIf OB_NO9.Value = True Then
            ActiveCell.Offset(0, 52) = 0
            ActiveCell.Offset(0, 53) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON LA RAMPA EN ACCESOS CON PENDIENTES DE 6% CON BARANDAL")
        End If
        
        If OB_SI10.Value = True Then
            ActiveCell.Offset(0, 54) = 1
            ActiveCell.Offset(0, 55) = 0
        ElseIf OB_NO10.Value = True Then
            ActiveCell.Offset(0, 54) = 0
            ActiveCell.Offset(0, 55) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON LA RAMPA EN ACCESOS CON PENDIENTES DE 6% CON BARANDAL")
        End If
        
        If OB_SI11.Value = True Then
            ActiveCell.Offset(0, 56) = 1
            ActiveCell.Offset(0, 57) = 0
        ElseIf OB_NO11.Value = True Then
            ActiveCell.Offset(0, 56) = 0
            ActiveCell.Offset(0, 57) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON PUERTAS DE ACCESO DE 1.20 MTS. ANCHO CON PUERTA ABATIBLE")
        End If
        
        If OB_SI12.Value = True Then
            ActiveCell.Offset(0, 58) = 1
            ActiveCell.Offset(0, 59) = 0
        ElseIf OB_NO12.Value = True Then
            ActiveCell.Offset(0, 58) = 0
            ActiveCell.Offset(0, 59) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON PUERTAS DE ACCESO DE 1.20 MTS. ANCHO CON PUERTA ABATIBLE")
        End If
        
        If OB_SI13.Value = True Then
            ActiveCell.Offset(0, 60) = 1
            ActiveCell.Offset(0, 61) = 0
        ElseIf OB_NO13.Value = True Then
            ActiveCell.Offset(0, 60) = 0
            ActiveCell.Offset(0, 61) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON LA RAMPA EN ACCESOS CON PENDIENTES DE 6% CON PISO ANTIDERRAPANTE")
        End If
        
        If OB_SI14.Value = True Then
            ActiveCell.Offset(0, 62) = 1
            ActiveCell.Offset(0, 63) = 0
        ElseIf OB_NO14.Value = True Then
            ActiveCell.Offset(0, 62) = 0
            ActiveCell.Offset(0, 63) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON LA RAMPA EN ACCESOS CON PENDIENTES DE 6% CON PISO ANTIDERRAPANTE")
        End If
        
        If OB_SI15.Value = True Then
            ActiveCell.Offset(0, 64) = 1
            ActiveCell.Offset(0, 65) = 0
        ElseIf OB_NO15.Value = True Then
            ActiveCell.Offset(0, 64) = 0
            ActiveCell.Offset(0, 65) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON PUERTAS DE ACCESO DE 1.20 MTS. ANCHO CON PUERTAS CORREDIZAS")
        End If
        
        If OB_SI16.Value = True Then
            ActiveCell.Offset(0, 66) = 1
            ActiveCell.Offset(0, 67) = 0
        ElseIf OB_NO16.Value = True Then
            ActiveCell.Offset(0, 66) = 0
            ActiveCell.Offset(0, 67) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON PUERTAS DE ACCESO DE 1.20 MTS. ANCHO CON PUERTAS CORREDIZAS")
        End If
        
        ActiveCell.Offset(0, 68) = Observaciones1
        ActiveCell.Offset(0, 69) = Observaciones2
        ActiveCell.Offset(0, 70) = Observaciones3
        ActiveCell.Offset(0, 71) = Observaciones4
        
        'VESTÍBULO
        
        If OB_SI17.Value = True Then
            ActiveCell.Offset(0, 72) = 1
            ActiveCell.Offset(0, 73) = 0
        ElseIf OB_NO17.Value = True Then
            ActiveCell.Offset(0, 72) = 0
            ActiveCell.Offset(0, 73) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON SILLA DE RUEDAS EN ACCESO")
        End If
        
        If OB_SI18.Value = True Then
            ActiveCell.Offset(0, 74) = 1
            ActiveCell.Offset(0, 75) = 0
        ElseIf OB_NO18.Value = True Then
            ActiveCell.Offset(0, 74) = 0
            ActiveCell.Offset(0, 75) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON SILLA DE RUEDAS EN ACCESO")
        End If
        
        If OB_SI19.Value = True Then
            ActiveCell.Offset(0, 76) = 1
            ActiveCell.Offset(0, 77) = 0
        ElseIf OB_NO19.Value = True Then
            ActiveCell.Offset(0, 76) = 0
            ActiveCell.Offset(0, 77) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON MÓDULO DE INFORMACIÓN Y ATENCIÓN AL PÚBLICO")
        End If
        
        If OB_SI20.Value = True Then
            ActiveCell.Offset(0, 78) = 1
            ActiveCell.Offset(0, 79) = 0
        ElseIf OB_NO20.Value = True Then
            ActiveCell.Offset(0, 78) = 0
            ActiveCell.Offset(0, 79) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON MÓDULO DE INFORMACIÓN Y ATENCIÓN AL PÚBLICO")
        End If
        
        If OB_SI21.Value = True Then
            ActiveCell.Offset(0, 80) = 1
            ActiveCell.Offset(0, 81) = 0
        ElseIf OB_NO21.Value = True Then
            ActiveCell.Offset(0, 80) = 0
            ActiveCell.Offset(0, 81) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON ELEVADOR EN VESTÍBULO")
        End If
        
        If OB_SI22.Value = True Then
            ActiveCell.Offset(0, 82) = 1
            ActiveCell.Offset(0, 83) = 0
        ElseIf OB_NO22.Value = True Then
            ActiveCell.Offset(0, 82) = 0
            ActiveCell.Offset(0, 83) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON ELEVADOR EN VESTÍBULO")
        End If
        
        If OB_SI23.Value = True Then
            ActiveCell.Offset(0, 84) = 1
            ActiveCell.Offset(0, 85) = 0
        ElseIf OB_NO23.Value = True Then
            ActiveCell.Offset(0, 84) = 0
            ActiveCell.Offset(0, 85) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON DIRECTORIO EN VESTÍBULO")
        End If
        
        If OB_SI24.Value = True Then
            ActiveCell.Offset(0, 86) = 1
            ActiveCell.Offset(0, 87) = 0
        ElseIf OB_NO24.Value = True Then
            ActiveCell.Offset(0, 86) = 0
            ActiveCell.Offset(0, 87) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON DIRECTORIO EN VESTÍBULO")
        End If
        
        If OB_SI25.Value = True Then
            ActiveCell.Offset(0, 88) = 1
            ActiveCell.Offset(0, 89) = 0
        ElseIf OB_NO25.Value = True Then
            ActiveCell.Offset(0, 88) = 0
            ActiveCell.Offset(0, 89) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON FOLLETOS DE INFORMACIÓN")
        End If
        
        If OB_SI26.Value = True Then
            ActiveCell.Offset(0, 90) = 1
            ActiveCell.Offset(0, 91) = 0
        ElseIf OB_NO26.Value = True Then
            ActiveCell.Offset(0, 90) = 0
            ActiveCell.Offset(0, 91) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON FOLLETOS DE INFORMACIÓN")
        End If
        
        If OB_SI27.Value = True Then
            ActiveCell.Offset(0, 92) = 1
            ActiveCell.Offset(0, 93) = 0
        ElseIf OB_NO27.Value = True Then
            ActiveCell.Offset(0, 92) = 0
            ActiveCell.Offset(0, 93) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON SANITARIOS EN VESTÍBULO")
        End If
        
        If OB_SI28.Value = True Then
            ActiveCell.Offset(0, 94) = 1
            ActiveCell.Offset(0, 95) = 0
        ElseIf OB_NO28.Value = True Then
            ActiveCell.Offset(0, 94) = 0
            ActiveCell.Offset(0, 95) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON SANITARIOS EN VESTÍBULO")
        End If
        
        If OB_SI29.Value = True Then
            ActiveCell.Offset(0, 96) = 1
            ActiveCell.Offset(0, 97) = 0
        ElseIf OB_NO29.Value = True Then
            ActiveCell.Offset(0, 96) = 0
            ActiveCell.Offset(0, 97) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON UN MÓDULO DE INFORMACIÓN Y ATENCIÓN AL PÚBLICO A UNA ALTURA DE 0.80 MTS.")
        End If
        
        If OB_SI30.Value = True Then
            ActiveCell.Offset(0, 98) = 1
            ActiveCell.Offset(0, 99) = 0
        ElseIf OB_NO30.Value = True Then
            ActiveCell.Offset(0, 98) = 0
            ActiveCell.Offset(0, 99) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON UN MÓDULO DE INFORMACIÓN Y ATENCIÓN AL PÚBLICO A UNA ALTURA DE 0.80 MTS.")
        End If
        
        If OB_SI31.Value = True Then
            ActiveCell.Offset(0, 100) = 1
            ActiveCell.Offset(0, 101) = 0
        ElseIf OB_NO31.Value = True Then
            ActiveCell.Offset(0, 100) = 0
            ActiveCell.Offset(0, 101) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON ELEVADOR EN VESTÍBULO CON BOTONES BRAILLE")
        End If
        
        If OB_SI32.Value = True Then
            ActiveCell.Offset(0, 102) = 1
            ActiveCell.Offset(0, 103) = 0
        ElseIf OB_NO32.Value = True Then
            ActiveCell.Offset(0, 102) = 0
            ActiveCell.Offset(0, 103) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON ELEVADOR EN VESTÍBULO CON BOTONES BRAILLE")
        End If
        
        If OB_SI33.Value = True Then
            ActiveCell.Offset(0, 104) = 1
            ActiveCell.Offset(0, 105) = 0
        ElseIf OB_NO33.Value = True Then
            ActiveCell.Offset(0, 104) = 0
            ActiveCell.Offset(0, 105) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON DIRECTORIO EN VESTÍBULO EN SISTEMA BRAILLE")
        End If
        
        If OB_SI34.Value = True Then
            ActiveCell.Offset(0, 106) = 1
            ActiveCell.Offset(0, 107) = 0
        ElseIf OB_NO34.Value = True Then
            ActiveCell.Offset(0, 106) = 0
            ActiveCell.Offset(0, 107) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON DIRECTORIO EN VESTÍBULO EN SISTEMA BRAILLE")
        End If
        
        If OB_SI35.Value = True Then
            ActiveCell.Offset(0, 108) = 1
            ActiveCell.Offset(0, 109) = 0
        ElseIf OB_NO35.Value = True Then
            ActiveCell.Offset(0, 108) = 0
            ActiveCell.Offset(0, 109) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON FOLLETOS DE INFORMACIÓN EN SISTEMA BRAILLE")
        End If
        
        If OB_SI36.Value = True Then
            ActiveCell.Offset(0, 110) = 1
            ActiveCell.Offset(0, 111) = 0
        ElseIf OB_NO36.Value = True Then
            ActiveCell.Offset(0, 110) = 0
            ActiveCell.Offset(0, 111) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON FOLLETOS DE INFORMACIÓN EN SISTEMA BRAILLE")
        End If
        
        If OB_SI37.Value = True Then
            ActiveCell.Offset(0, 112) = 1
            ActiveCell.Offset(0, 113) = 0
        ElseIf OB_NO37.Value = True Then
            ActiveCell.Offset(0, 112) = 0
            ActiveCell.Offset(0, 113) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON SANITARIOS EN VESTÍBULO PARA DISCAPACITADOS")
        End If
        
        If OB_SI38.Value = True Then
            ActiveCell.Offset(0, 114) = 1
            ActiveCell.Offset(0, 115) = 0
        ElseIf OB_NO38.Value = True Then
            ActiveCell.Offset(0, 114) = 0
            ActiveCell.Offset(0, 115) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON SANITARIOS EN VESTÍBULO PARA DISCAPACITADOS")
        End If
        
        If OB_SI39.Value = True Then
            ActiveCell.Offset(0, 116) = 1
            ActiveCell.Offset(0, 117) = 0
        ElseIf OB_NO39.Value = True Then
            ActiveCell.Offset(0, 116) = 0
            ActiveCell.Offset(0, 117) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON UN MÓDULO DE INFORMACIÓN Y ATENCIÓN AL PÚBLICO CON PERSONAL INTERPRETE")
        End If
        
        If OB_SI40.Value = True Then
            ActiveCell.Offset(0, 118) = 1
            ActiveCell.Offset(0, 119) = 0
        ElseIf OB_NO40.Value = True Then
            ActiveCell.Offset(0, 118) = 0
            ActiveCell.Offset(0, 119) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON UN MÓDULO DE INFORMACIÓN Y ATENCIÓN AL PÚBLICO CON PERSONAL INTERPRETE")
        End If
        
        If OB_SI41.Value = True Then
            ActiveCell.Offset(0, 120) = 1
            ActiveCell.Offset(0, 121) = 0
        ElseIf OB_NO41.Value = True Then
            ActiveCell.Offset(0, 120) = 0
            ActiveCell.Offset(0, 121) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON UN ELEVADOR EN VESTÍBULO A UNA ALTURA DE 0.90 MTS.")
        End If
        
        If OB_SI42.Value = True Then
            ActiveCell.Offset(0, 122) = 1
            ActiveCell.Offset(0, 123) = 0
        ElseIf OB_NO42.Value = True Then
            ActiveCell.Offset(0, 122) = 0
            ActiveCell.Offset(0, 123) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON UN ELEVADOR EN VESTÍBULO A UNA ALTURA DE 0.90 MTS.")
        End If
        
        If OB_SI43.Value = True Then
            ActiveCell.Offset(0, 124) = 1
            ActiveCell.Offset(0, 125) = 0
        ElseIf OB_NO43.Value = True Then
            ActiveCell.Offset(0, 124) = 0
            ActiveCell.Offset(0, 125) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON UN DIRECTORIO EN VESTÍBULO A UNA ALTURA DE 0.90 MTS.")
        End If
        
        If OB_SI44.Value = True Then
            ActiveCell.Offset(0, 126) = 1
            ActiveCell.Offset(0, 127) = 0
        ElseIf OB_NO44.Value = True Then
            ActiveCell.Offset(0, 126) = 0
            ActiveCell.Offset(0, 127) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON UN DIRECTORIO EN VESTÍBULO A UNA ALTURA DE 0.90 MTS.")
        End If
        
        ActiveCell.Offset(0, 128) = Observaciones5
        ActiveCell.Offset(0, 129) = Observaciones6
        ActiveCell.Offset(0, 130) = Observaciones7
        ActiveCell.Offset(0, 131) = Observaciones8
        ActiveCell.Offset(0, 132) = Observaciones9
        ActiveCell.Offset(0, 133) = Observaciones10
        
        'Circulaciones
        
        If OB_SI45.Value = True Then
            ActiveCell.Offset(0, 134) = 1
            ActiveCell.Offset(0, 135) = 0
        ElseIf OB_NO45.Value = True Then
            ActiveCell.Offset(0, 134) = 0
            ActiveCell.Offset(0, 135) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON PRINCIPALES DE 1.20 MTS. DE ANCHO MÍNIMO")
        End If
        
        If OB_SI46.Value = True Then
            ActiveCell.Offset(0, 136) = 1
            ActiveCell.Offset(0, 137) = 0
        ElseIf OB_NO46.Value = True Then
            ActiveCell.Offset(0, 136) = 0
            ActiveCell.Offset(0, 137) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON PRINCIPALES DE 1.20 MTS. DE ANCHO MÍNIMO")
        End If
        
        If OB_SI47.Value = True Then
            ActiveCell.Offset(0, 138) = 1
            ActiveCell.Offset(0, 139) = 0
        ElseIf OB_NO47.Value = True Then
            ActiveCell.Offset(0, 138) = 0
            ActiveCell.Offset(0, 139) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON INTERNAS DE 0.90 MTS. DE ANCHO MÍNIMO")
        End If
        
        If OB_SI48.Value = True Then
            ActiveCell.Offset(0, 140) = 1
            ActiveCell.Offset(0, 141) = 0
        ElseIf OB_NO48.Value = True Then
            ActiveCell.Offset(0, 140) = 0
            ActiveCell.Offset(0, 141) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON INTERNAS DE 0.90 MTS. DE ANCHO MÍNIMO")
        End If
        
        If OB_SI49.Value = True Then
            ActiveCell.Offset(0, 142) = 1
            ActiveCell.Offset(0, 143) = 0
        ElseIf OB_NO49.Value = True Then
            ActiveCell.Offset(0, 142) = 0
            ActiveCell.Offset(0, 143) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON CIRCULACIONES LIBRES DE OBSTÁCULOS FIJOS")
        End If
        
        If OB_SI50.Value = True Then
            ActiveCell.Offset(0, 144) = 1
            ActiveCell.Offset(0, 145) = 0
        ElseIf OB_NO50.Value = True Then
            ActiveCell.Offset(0, 144) = 0
            ActiveCell.Offset(0, 145) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON CIRCULACIONES LIBRES DE OBSTÁCULOS FIJOS")
        End If
        
        If OB_SI51.Value = True Then
            ActiveCell.Offset(0, 146) = 1
            ActiveCell.Offset(0, 147) = 0
        ElseIf OB_NO51.Value = True Then
            ActiveCell.Offset(0, 146) = 0
            ActiveCell.Offset(0, 147) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON PRINCIPALES DE 1.20 MTS. DE ANCHO MÍNIMO CON BARANDALES")
        End If
        
        If OB_SI52.Value = True Then
            ActiveCell.Offset(0, 148) = 1
            ActiveCell.Offset(0, 149) = 0
        ElseIf OB_NO52.Value = True Then
            ActiveCell.Offset(0, 148) = 0
            ActiveCell.Offset(0, 149) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON PRINCIPALES DE 1.20 MTS. DE ANCHO MÍNIMO CON BARANDALES")
        End If
                
        If OB_SI53.Value = True Then
            ActiveCell.Offset(0, 150) = 1
            ActiveCell.Offset(0, 151) = 0
        ElseIf OB_NO53.Value = True Then
            ActiveCell.Offset(0, 150) = 0
            ActiveCell.Offset(0, 151) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON INTERNAS DE 0.90 MTS. DE ANCHO MÍNIMO CON BARANDAL")
        End If
        
        If OB_SI54.Value = True Then
            ActiveCell.Offset(0, 152) = 1
            ActiveCell.Offset(0, 153) = 0
        ElseIf OB_NO54.Value = True Then
            ActiveCell.Offset(0, 152) = 0
            ActiveCell.Offset(0, 153) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON INTERNAS DE 0.90 MTS. DE ANCHO MÍNIMO CON BARANDAL")
        End If
        
        If OB_SI55.Value = True Then
            ActiveCell.Offset(0, 154) = 1
            ActiveCell.Offset(0, 155) = 0
        ElseIf OB_NO55.Value = True Then
            ActiveCell.Offset(0, 154) = 0
            ActiveCell.Offset(0, 155) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON CIRCULACIONES LIBRES DE OBSTÁCULOS MOVILES")
        End If
        
        If OB_SI56.Value = True Then
            ActiveCell.Offset(0, 156) = 1
            ActiveCell.Offset(0, 157) = 0
        ElseIf OB_NO56.Value = True Then
            ActiveCell.Offset(0, 156) = 0
            ActiveCell.Offset(0, 157) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON CIRCULACIONES LIBRES DE OBSTÁCULOS MOVILES")
        End If
        
        If OB_SI57.Value = True Then
            ActiveCell.Offset(0, 158) = 1
            ActiveCell.Offset(0, 159) = 0
        ElseIf OB_NO57.Value = True Then
            ActiveCell.Offset(0, 158) = 0
            ActiveCell.Offset(0, 159) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON PRINCIPALES DE 1.20 MTS. DE ANCHO MÍNIMO CON PISO ANTIDERRAPANTE")
        End If
        
        If OB_SI58.Value = True Then
            ActiveCell.Offset(0, 160) = 1
            ActiveCell.Offset(0, 161) = 0
        ElseIf OB_NO58.Value = True Then
            ActiveCell.Offset(0, 160) = 0
            ActiveCell.Offset(0, 161) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON PRINCIPALES DE 1.20 MTS. DE ANCHO MÍNIMO CON PISO ANTIDERRAPANTE")
        End If
        
        If OB_SI59.Value = True Then
            ActiveCell.Offset(0, 162) = 1
            ActiveCell.Offset(0, 163) = 0
        ElseIf OB_NO59.Value = True Then
            ActiveCell.Offset(0, 162) = 0
            ActiveCell.Offset(0, 163) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON INTERNAS DE 0.90 MTS. DE ANCHO MÍNIMO CON PISO ANTIDERRAPANTE")
        End If
        
        If OB_SI60.Value = True Then
            ActiveCell.Offset(0, 164) = 1
            ActiveCell.Offset(0, 165) = 0
        ElseIf OB_NO60.Value = True Then
            ActiveCell.Offset(0, 164) = 0
            ActiveCell.Offset(0, 165) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON INTERNAS DE 0.90 MTS. DE ANCHO MÍNIMO CON PISO ANTIDERRAPANTE")
        End If
        
        If OB_SI61.Value = True Then
            ActiveCell.Offset(0, 166) = 1
            ActiveCell.Offset(0, 167) = 0
        ElseIf OB_NO61.Value = True Then
            ActiveCell.Offset(0, 166) = 0
            ActiveCell.Offset(0, 167) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON CIRCULACIONES LIBRES CON BUENA ILUMINACIÓN")
        End If
        
        If OB_SI62.Value = True Then
            ActiveCell.Offset(0, 168) = 1
            ActiveCell.Offset(0, 169) = 0
        ElseIf OB_NO62.Value = True Then
            ActiveCell.Offset(0, 168) = 0
            ActiveCell.Offset(0, 169) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON CIRCULACIONES LIBRES CON BUENA ILUMINACIÓN")
        End If
        
        ActiveCell.Offset(0, 170) = Observaciones11
        ActiveCell.Offset(0, 171) = Observaciones12
        ActiveCell.Offset(0, 172) = Observaciones13
        
        'Señalización
        
        If OB_SI63.Value = True Then
            ActiveCell.Offset(0, 173) = 1
            ActiveCell.Offset(0, 174) = 0
        ElseIf OB_NO63.Value = True Then
            ActiveCell.Offset(0, 173) = 0
            ActiveCell.Offset(0, 174) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON VISUAL EN EL ACCESO")
        End If
        
        If OB_SI64.Value = True Then
            ActiveCell.Offset(0, 175) = 1
            ActiveCell.Offset(0, 176) = 0
        ElseIf OB_NO64.Value = True Then
            ActiveCell.Offset(0, 175) = 0
            ActiveCell.Offset(0, 176) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON VISUAL EN EL ACCESO")
        End If
        
        If OB_SI65.Value = True Then
            ActiveCell.Offset(0, 177) = 1
            ActiveCell.Offset(0, 178) = 0
        ElseIf OB_NO65.Value = True Then
            ActiveCell.Offset(0, 177) = 0
            ActiveCell.Offset(0, 178) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON SISTEMA BRAILLE EN EL ACCESO")
        End If
        
        If OB_SI66.Value = True Then
            ActiveCell.Offset(0, 179) = 1
            ActiveCell.Offset(0, 180) = 0
        ElseIf OB_NO66.Value = True Then
            ActiveCell.Offset(0, 179) = 0
            ActiveCell.Offset(0, 180) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON SISTEMA BRAILLE EN EL ACCESO")
        End If
                
        If OB_SI67.Value = True Then
            ActiveCell.Offset(0, 181) = 1
            ActiveCell.Offset(0, 182) = 0
        ElseIf OB_NO67.Value = True Then
            ActiveCell.Offset(0, 181) = 0
            ActiveCell.Offset(0, 182) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON VISUAL Y TÁCTIL EN PISO")
        End If
        
        If OB_SI68.Value = True Then
            ActiveCell.Offset(0, 183) = 1
            ActiveCell.Offset(0, 184) = 0
        ElseIf OB_NO68.Value = True Then
            ActiveCell.Offset(0, 183) = 0
            ActiveCell.Offset(0, 184) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON VISUAL Y TÁCTIL EN PISO")
        End If
    
        If OB_SI69.Value = True Then
            ActiveCell.Offset(0, 185) = 1
            ActiveCell.Offset(0, 186) = 0
        ElseIf OB_NO69.Value = True Then
            ActiveCell.Offset(0, 185) = 0
            ActiveCell.Offset(0, 186) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON VISUAL EN EL VESTÍBULO")
        End If
        
        If OB_SI70.Value = True Then
            ActiveCell.Offset(0, 187) = 1
            ActiveCell.Offset(0, 188) = 0
        ElseIf OB_NO70.Value = True Then
            ActiveCell.Offset(0, 187) = 0
            ActiveCell.Offset(0, 188) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON VISUAL EN EL VESTÍBULO")
        End If
        
        If OB_SI71.Value = True Then
            ActiveCell.Offset(0, 189) = 1
            ActiveCell.Offset(0, 190) = 0
        ElseIf OB_NO71.Value = True Then
            ActiveCell.Offset(0, 189) = 0
            ActiveCell.Offset(0, 190) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON SISTEMA BRAILLE EN EL VESTÍBULO")
        End If
        
        If OB_SI72.Value = True Then
            ActiveCell.Offset(0, 191) = 1
            ActiveCell.Offset(0, 192) = 0
        ElseIf OB_NO72.Value = True Then
            ActiveCell.Offset(0, 191) = 0
            ActiveCell.Offset(0, 192) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON SISTEMA BRAILLE EN EL VESTÍBULO")
        End If
                
        If OB_SI73.Value = True Then
            ActiveCell.Offset(0, 193) = 1
            ActiveCell.Offset(0, 194) = 0
        ElseIf OB_NO73.Value = True Then
            ActiveCell.Offset(0, 193) = 0
            ActiveCell.Offset(0, 194) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON VISUAL Y TÁCTIL EN PISO DEL VESTÍBULO")
        End If
        
        If OB_SI74.Value = True Then
            ActiveCell.Offset(0, 195) = 1
            ActiveCell.Offset(0, 196) = 0
        ElseIf OB_NO74.Value = True Then
            ActiveCell.Offset(0, 195) = 0
            ActiveCell.Offset(0, 196) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON VISUAL Y TÁCTIL EN PISO DEL VESTÍBULO")
        End If
      
        If OB_SI75.Value = True Then
            ActiveCell.Offset(0, 197) = 1
            ActiveCell.Offset(0, 198) = 0
        ElseIf OB_NO75.Value = True Then
            ActiveCell.Offset(0, 197) = 0
            ActiveCell.Offset(0, 198) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON VISUAL EN CIRCULACIONES")
        End If
        
        If OB_SI76.Value = True Then
            ActiveCell.Offset(0, 199) = 1
            ActiveCell.Offset(0, 200) = 0
        ElseIf OB_NO76.Value = True Then
            ActiveCell.Offset(0, 199) = 0
            ActiveCell.Offset(0, 200) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON VISUAL EN CIRCULACIONES")
        End If
        
        If OB_SI77.Value = True Then
            ActiveCell.Offset(0, 201) = 1
            ActiveCell.Offset(0, 202) = 0
        ElseIf OB_NO77.Value = True Then
            ActiveCell.Offset(0, 201) = 0
            ActiveCell.Offset(0, 202) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON SISTEMA BRAILLE EN CIRCULACIONES")
        End If
        
        If OB_SI78.Value = True Then
            ActiveCell.Offset(0, 203) = 1
            ActiveCell.Offset(0, 204) = 0
        ElseIf OB_NO78.Value = True Then
            ActiveCell.Offset(0, 203) = 0
            ActiveCell.Offset(0, 204) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON SISTEMA BRAILLE EN CIRCULACIONES")
        End If
                
        If OB_SI79.Value = True Then
            ActiveCell.Offset(0, 205) = 1
            ActiveCell.Offset(0, 206) = 0
        ElseIf OB_NO79.Value = True Then
            ActiveCell.Offset(0, 205) = 0
            ActiveCell.Offset(0, 206) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON VISUAL Y TÁCTIL EN CIRCULACIONES")
        End If
        
        If OB_SI80.Value = True Then
            ActiveCell.Offset(0, 207) = 1
            ActiveCell.Offset(0, 208) = 0
        ElseIf OB_NO80.Value = True Then
            ActiveCell.Offset(0, 207) = 0
            ActiveCell.Offset(0, 208) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON VISUAL Y TÁCTIL EN CIRCULACIONES")
        End If
        
        ActiveCell.Offset(0, 209) = Observaciones14
        ActiveCell.Offset(0, 210) = Observaciones15
        ActiveCell.Offset(0, 211) = Observaciones16
        
        'Edificio y servicio
                
        If OB_SI81.Value = True Then
            ActiveCell.Offset(0, 212) = 1
            ActiveCell.Offset(0, 213) = 0
        ElseIf OB_NO81.Value = True Then
            ActiveCell.Offset(0, 212) = 0
            ActiveCell.Offset(0, 213) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON ACCESO A TODOS LOS NIVELES")
        End If
        
        If OB_SI82.Value = True Then
            ActiveCell.Offset(0, 214) = 1
            ActiveCell.Offset(0, 215) = 0
        ElseIf OB_NO82.Value = True Then
            ActiveCell.Offset(0, 214) = 0
            ActiveCell.Offset(0, 215) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON ACCESO A TODOS LOS NIVELES")
        End If
        
        If OB_SI83.Value = True Then
            ActiveCell.Offset(0, 216) = 1
            ActiveCell.Offset(0, 217) = 0
        ElseIf OB_NO83.Value = True Then
            ActiveCell.Offset(0, 216) = 0
            ActiveCell.Offset(0, 217) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON ALARMAS VISUALES EN TODOS LOS NIVELES")
        End If
        
        If OB_SI84.Value = True Then
            ActiveCell.Offset(0, 218) = 1
            ActiveCell.Offset(0, 219) = 0
        ElseIf OB_NO84.Value = True Then
            ActiveCell.Offset(0, 218) = 0
            ActiveCell.Offset(0, 219) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON ALARMAS VISUALES EN TODOS LOS NIVELES")
        End If
        
        If OB_SI85.Value = True Then
            ActiveCell.Offset(0, 220) = 1
            ActiveCell.Offset(0, 221) = 0
        ElseIf OB_NO85.Value = True Then
            ActiveCell.Offset(0, 220) = 0
            ActiveCell.Offset(0, 221) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON ALARMAS AUDITIVAS EN TODOS LOS NIVELES")
        End If
        
        If OB_SI86.Value = True Then
            ActiveCell.Offset(0, 222) = 1
            ActiveCell.Offset(0, 223) = 0
        ElseIf OB_NO86.Value = True Then
            ActiveCell.Offset(0, 222) = 0
            ActiveCell.Offset(0, 223) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON ALARMAS AUDITIVAS EN TODOS LOS NIVELES")
        End If
        
        If OB_SI87.Value = True Then
            ActiveCell.Offset(0, 224) = 1
            ActiveCell.Offset(0, 225) = 0
        ElseIf OB_NO87.Value = True Then
            ActiveCell.Offset(0, 224) = 0
            ActiveCell.Offset(0, 225) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON RUTA DE ACCESIBILIDAD HACIA TODOS LOS SERVICIOS")
        End If
        
        If OB_SI88.Value = True Then
            ActiveCell.Offset(0, 226) = 1
            ActiveCell.Offset(0, 227) = 0
        ElseIf OB_NO88.Value = True Then
            ActiveCell.Offset(0, 226) = 0
            ActiveCell.Offset(0, 227) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON RUTA DE ACCESIBILIDAD HACIA TODOS LOS SERVICIOS")
        End If
        
        If OB_SI89.Value = True Then
            ActiveCell.Offset(0, 228) = 1
            ActiveCell.Offset(0, 229) = 0
        ElseIf OB_NO89.Value = True Then
            ActiveCell.Offset(0, 228) = 0
            ActiveCell.Offset(0, 229) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON ACCESO A ÁREAS PÚBLICAS")
        End If
        
        If OB_SI90.Value = True Then
            ActiveCell.Offset(0, 230) = 1
            ActiveCell.Offset(0, 231) = 0
        ElseIf OB_NO90.Value = True Then
            ActiveCell.Offset(0, 230) = 0
            ActiveCell.Offset(0, 231) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON ACCESO A ÁREAS PÚBLICAS")
        End If
        
        If OB_SI91.Value = True Then
            ActiveCell.Offset(0, 232) = 1
            ActiveCell.Offset(0, 233) = 0
        ElseIf OB_NO91.Value = True Then
            ActiveCell.Offset(0, 232) = 0
            ActiveCell.Offset(0, 233) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON ALARMAS VISUALES EN ÁREAS PÚBLICAS")
        End If
        
        If OB_SI92.Value = True Then
            ActiveCell.Offset(0, 234) = 1
            ActiveCell.Offset(0, 235) = 0
        ElseIf OB_NO92.Value = True Then
            ActiveCell.Offset(0, 234) = 0
            ActiveCell.Offset(0, 235) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON ALARMAS VISUALES EN ÁREAS PÚBLICAS")
        End If
        
        If OB_SI93.Value = True Then
            ActiveCell.Offset(0, 236) = 1
            ActiveCell.Offset(0, 237) = 0
        ElseIf OB_NO93.Value = True Then
            ActiveCell.Offset(0, 236) = 0
            ActiveCell.Offset(0, 237) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON ALARMAS AUDITIVAS EN ÁREAS PÚBLICAS")
        End If
        
        If OB_SI94.Value = True Then
            ActiveCell.Offset(0, 238) = 1
            ActiveCell.Offset(0, 239) = 0
        ElseIf OB_NO94.Value = True Then
            ActiveCell.Offset(0, 238) = 0
            ActiveCell.Offset(0, 239) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON ALARMAS AUDITIVAS EN ÁREAS PÚBLICAS")
        End If
        
        If OB_SI95.Value = True Then
            ActiveCell.Offset(0, 240) = 1
            ActiveCell.Offset(0, 241) = 0
        ElseIf OB_NO95.Value = True Then
            ActiveCell.Offset(0, 240) = 0
            ActiveCell.Offset(0, 241) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON RUTA DE ACCESIBILIDAD HACIA ÁREAS PÚBLICAS")
        End If
        
        If OB_SI96.Value = True Then
            ActiveCell.Offset(0, 242) = 1
            ActiveCell.Offset(0, 243) = 0
        ElseIf OB_NO96.Value = True Then
            ActiveCell.Offset(0, 242) = 0
            ActiveCell.Offset(0, 243) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON RUTA DE ACCESIBILIDAD HACIA ÁREAS PÚBLICAS")
        End If
        
        If OB_SI97.Value = True Then
            ActiveCell.Offset(0, 244) = 1
            ActiveCell.Offset(0, 245) = 0
        ElseIf OB_NO97.Value = True Then
            ActiveCell.Offset(0, 244) = 0
            ActiveCell.Offset(0, 245) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON ACCESO A ÁREAS PRIVADAS")
        End If
        
        If OB_SI98.Value = True Then
            ActiveCell.Offset(0, 246) = 1
            ActiveCell.Offset(0, 247) = 0
        ElseIf OB_NO98.Value = True Then
            ActiveCell.Offset(0, 246) = 0
            ActiveCell.Offset(0, 247) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON ACCESO A ÁREAS PRIVADAS")
        End If
        
        If OB_SI99.Value = True Then
            ActiveCell.Offset(0, 248) = 1
            ActiveCell.Offset(0, 249) = 0
        ElseIf OB_NO99.Value = True Then
            ActiveCell.Offset(0, 248) = 0
            ActiveCell.Offset(0, 249) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON ALARMAS VISUALES EN ÁREAS PRIVADAS")
        End If
        
        If OB_SI100.Value = True Then
            ActiveCell.Offset(0, 250) = 1
            ActiveCell.Offset(0, 251) = 0
        ElseIf OB_NO100.Value = True Then
            ActiveCell.Offset(0, 250) = 0
            ActiveCell.Offset(0, 251) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON ALARMAS VISUALES EN ÁREAS PRIVADAS")
        End If
        
        If OB_SI101.Value = True Then
            ActiveCell.Offset(0, 252) = 1
            ActiveCell.Offset(0, 253) = 0
        ElseIf OB_NO101.Value = True Then
            ActiveCell.Offset(0, 252) = 0
            ActiveCell.Offset(0, 253) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON ALARMAS AUDITIVAS EN ÁREAS PRIVADAS")
        End If
        
        If OB_SI102.Value = True Then
            ActiveCell.Offset(0, 254) = 1
            ActiveCell.Offset(0, 255) = 0
        ElseIf OB_NO102.Value = True Then
            ActiveCell.Offset(0, 254) = 0
            ActiveCell.Offset(0, 255) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON ALARMAS AUDITIVAS EN ÁREAS PRIVADAS")
        End If
        
        If OB_SI103.Value = True Then
            ActiveCell.Offset(0, 256) = 1
            ActiveCell.Offset(0, 257) = 0
        ElseIf OB_NO103.Value = True Then
            ActiveCell.Offset(0, 256) = 0
            ActiveCell.Offset(0, 257) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON RUTA DE ACCESIBILIDAD HACIA ÁREAS PRIVADAS")
        End If
        
        If OB_SI104.Value = True Then
            ActiveCell.Offset(0, 258) = 1
            ActiveCell.Offset(0, 259) = 0
        ElseIf OB_NO104.Value = True Then
            ActiveCell.Offset(0, 258) = 0
            ActiveCell.Offset(0, 259) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON RUTA DE ACCESIBILIDAD HACIA ÁREAS PRIVADAS")
        End If
        
        ActiveCell.Offset(0, 260) = Observaciones17
        ActiveCell.Offset(0, 261) = Observaciones18
        ActiveCell.Offset(0, 262) = Observaciones19
        ActiveCell.Offset(0, 263) = Observaciones20
                
        'Sanitarios para uso exclusivo
        
         If OB_SI105.Value = True Then
            ActiveCell.Offset(0, 264) = 1
            ActiveCell.Offset(0, 265) = 0
        ElseIf OB_NO105.Value = True Then
            ActiveCell.Offset(0, 264) = 0
            ActiveCell.Offset(0, 265) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON W.C. CON DIMENSIONES DE 1.70X1.70 MTS.")
        End If
        
        If OB_SI106.Value = True Then
            ActiveCell.Offset(0, 266) = 1
            ActiveCell.Offset(0, 267) = 0
        ElseIf OB_NO106.Value = True Then
            ActiveCell.Offset(0, 266) = 0
            ActiveCell.Offset(0, 267) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON W.C. CON DIMENSIONES DE 1.70X1.70 MTS.")
        End If
                
        If OB_SI107.Value = True Then
            ActiveCell.Offset(0, 268) = 1
            ActiveCell.Offset(0, 269) = 0
        ElseIf OB_NO107.Value = True Then
            ActiveCell.Offset(0, 268) = 0
            ActiveCell.Offset(0, 269) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON MINGITORIO CON DIMENSIONES DE 0.90X0.40 MTS.")
        End If
        
        If OB_SI108.Value = True Then
            ActiveCell.Offset(0, 270) = 1
            ActiveCell.Offset(0, 271) = 0
        ElseIf OB_NO108.Value = True Then
            ActiveCell.Offset(0, 270) = 0
            ActiveCell.Offset(0, 271) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON MINGITORIO CON DIMENSIONES DE 0.90X0.40 MTS.")
        End If
        
        If OB_SI109.Value = True Then
            ActiveCell.Offset(0, 272) = 1
            ActiveCell.Offset(0, 273) = 0
        ElseIf OB_NO109.Value = True Then
            ActiveCell.Offset(0, 272) = 0
            ActiveCell.Offset(0, 273) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON LAVABO Y ESPEJO DIMENSIONES 0.90X0.45 MTS.")
        End If
        
        If OB_SI110.Value = True Then
            ActiveCell.Offset(0, 274) = 1
            ActiveCell.Offset(0, 275) = 0
        ElseIf OB_NO110.Value = True Then
            ActiveCell.Offset(0, 274) = 0
            ActiveCell.Offset(0, 275) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON LAVABO Y ESPEJO DIMENSIONES 0.90X0.45 MTS.")
        End If
        
        If OB_SI111.Value = True Then
            ActiveCell.Offset(0, 276) = 1
            ActiveCell.Offset(0, 277) = 0
        ElseIf OB_NO111.Value = True Then
            ActiveCell.Offset(0, 276) = 0
            ActiveCell.Offset(0, 277) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO LAS BARRAS DE APOYO EN EL W.C. CON DIMENSIONES DE 1.70X1.70 MTS.")
        End If
        
        If OB_SI112.Value = True Then
            ActiveCell.Offset(0, 278) = 1
            ActiveCell.Offset(0, 279) = 0
        ElseIf OB_NO112.Value = True Then
            ActiveCell.Offset(0, 278) = 0
            ActiveCell.Offset(0, 279) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO LAS BARRAS DE APOYO EN EL W.C. CON DIMENSIONES DE 1.70X1.70 MTS.")
        End If
                
        If OB_SI113.Value = True Then
            ActiveCell.Offset(0, 280) = 1
            ActiveCell.Offset(0, 281) = 0
        ElseIf OB_NO113.Value = True Then
            ActiveCell.Offset(0, 280) = 0
            ActiveCell.Offset(0, 281) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO LAS BARRAS DE APOYO EN EL MINGITORIO CON DIMENSIONES DE 0.90X0.40 MTS.")
        End If
        
        If OB_SI114.Value = True Then
            ActiveCell.Offset(0, 282) = 1
            ActiveCell.Offset(0, 283) = 0
        ElseIf OB_NO114.Value = True Then
            ActiveCell.Offset(0, 282) = 0
            ActiveCell.Offset(0, 283) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO LAS BARRAS DE APOYO EN EL MINGITORIO CON DIMENSIONES DE 0.90X0.40 MTS.")
        End If
        
        If OB_SI115.Value = True Then
            ActiveCell.Offset(0, 284) = 1
            ActiveCell.Offset(0, 285) = 0
        ElseIf OB_NO115.Value = True Then
            ActiveCell.Offset(0, 284) = 0
            ActiveCell.Offset(0, 285) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO LAS BARRAS DE APOYO EN EL LAVABO Y ESPEJO DIMENSIONES 0.90X0.45 MTS.")
        End If
        
        If OB_SI116.Value = True Then
            ActiveCell.Offset(0, 286) = 1
            ActiveCell.Offset(0, 287) = 0
        ElseIf OB_NO116.Value = True Then
            ActiveCell.Offset(0, 286) = 0
            ActiveCell.Offset(0, 287) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO LAS BARRAS DE APOYO EN EL LAVABO Y ESPEJO DIMENSIONES 0.90X0.45 MTS.")
        End If
       
        If OB_SI117.Value = True Then
            ActiveCell.Offset(0, 288) = 1
            ActiveCell.Offset(0, 289) = 0
        ElseIf OB_NO117.Value = True Then
            ActiveCell.Offset(0, 288) = 0
            ActiveCell.Offset(0, 289) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON W.C. DE ALTURA 0.45 MTS. CON DIMENSIONES DE 1.70X1.70 MTS.")
        End If
        
        If OB_SI118.Value = True Then
            ActiveCell.Offset(0, 290) = 1
            ActiveCell.Offset(0, 291) = 0
        ElseIf OB_NO118.Value = True Then
            ActiveCell.Offset(0, 290) = 0
            ActiveCell.Offset(0, 291) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON W.C. DE ALTURA 0.45 MTS. CON DIMENSIONES DE 1.70X1.70 MTS.")
        End If
                
        If OB_SI119.Value = True Then
            ActiveCell.Offset(0, 292) = 1
            ActiveCell.Offset(0, 293) = 0
        ElseIf OB_NO119.Value = True Then
            ActiveCell.Offset(0, 292) = 0
            ActiveCell.Offset(0, 293) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON MINGITORIO DE ALTURA 0.30 MTS. CON DIMENSIONES DE 0.90X0.40 MTS.")
        End If
        
        If OB_SI120.Value = True Then
            ActiveCell.Offset(0, 294) = 1
            ActiveCell.Offset(0, 295) = 0
        ElseIf OB_NO120.Value = True Then
            ActiveCell.Offset(0, 294) = 0
            ActiveCell.Offset(0, 295) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON MINGITORIO DE ALTURA 0.30 MTS. CON DIMENSIONES DE 0.90X0.40 MTS.")
        End If
        
        ActiveCell.Offset(0, 296) = Observaciones21
        ActiveCell.Offset(0, 297) = Observaciones22
        ActiveCell.Offset(0, 298) = Observaciones23
      
        'Ruta de evacuación emergente
        
        If OB_SI121.Value = True Then
            ActiveCell.Offset(0, 299) = 1
            ActiveCell.Offset(0, 300) = 0
        ElseIf OB_NO121.Value = True Then
            ActiveCell.Offset(0, 299) = 0
            ActiveCell.Offset(0, 300) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON UNA RUTA DE EVACUACIÓN EMERGENTE")
        End If
        
        If OB_SI122.Value = True Then
            ActiveCell.Offset(0, 301) = 1
            ActiveCell.Offset(0, 302) = 0
        ElseIf OB_NO122.Value = True Then
            ActiveCell.Offset(0, 301) = 0
            ActiveCell.Offset(0, 302) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON UNA RUTA DE EVACUACIÓN EMERGENTE")
        End If
        
        If OB_SI123.Value = True Then
            ActiveCell.Offset(0, 303) = 1
            ActiveCell.Offset(0, 304) = 0
        ElseIf OB_NO123.Value = True Then
            ActiveCell.Offset(0, 303) = 0
            ActiveCell.Offset(0, 304) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON UNA PUERTA DE EMERGENCIA DE 1.20 MTS. DE ANCHO MÍNIMO")
        End If
        
        If OB_SI124.Value = True Then
            ActiveCell.Offset(0, 305) = 1
            ActiveCell.Offset(0, 306) = 0
        ElseIf OB_NO124.Value = True Then
            ActiveCell.Offset(0, 305) = 0
            ActiveCell.Offset(0, 306) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON UNA PUERTA DE EMERGENCIA DE 1.20 MTS. DE ANCHO MÍNIMO")
        End If
        
        If OB_SI125.Value = True Then
            ActiveCell.Offset(0, 307) = 1
            ActiveCell.Offset(0, 308) = 0
        ElseIf OB_NO125.Value = True Then
            ActiveCell.Offset(0, 307) = 0
            ActiveCell.Offset(0, 308) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON RAMPAS DE 6% DE PENDIENTE MÁXIMA")
        End If
        
        If OB_SI126.Value = True Then
            ActiveCell.Offset(0, 309) = 1
            ActiveCell.Offset(0, 310) = 0
        ElseIf OB_NO126.Value = True Then
            ActiveCell.Offset(0, 309) = 0
            ActiveCell.Offset(0, 310) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON RAMPAS DE 6% DE PENDIENTE MÁXIMA")
        End If
                
        If OB_SI127.Value = True Then
            ActiveCell.Offset(0, 311) = 1
            ActiveCell.Offset(0, 312) = 0
        ElseIf OB_NO127.Value = True Then
            ActiveCell.Offset(0, 311) = 0
            ActiveCell.Offset(0, 312) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON SEÑALIZACIÓN EN LA RUTA DE EVACUACIÓN EMERGENTE")
        End If
        
        If OB_SI128.Value = True Then
            ActiveCell.Offset(0, 313) = 1
            ActiveCell.Offset(0, 314) = 0
        ElseIf OB_NO128.Value = True Then
            ActiveCell.Offset(0, 313) = 0
            ActiveCell.Offset(0, 314) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON SEÑALIZACIÓN EN LA RUTA DE EVACUACIÓN EMERGENTE")
        End If
        
        If OB_SI129.Value = True Then
            ActiveCell.Offset(0, 315) = 1
            ActiveCell.Offset(0, 316) = 0
        ElseIf OB_NO129.Value = True Then
            ActiveCell.Offset(0, 315) = 0
            ActiveCell.Offset(0, 316) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON UNA PUERTA DE EMERGENCIA DE 1.20 MTS. DE ANCHO MÍNIMO HACIA EL EXTERIOR")
        End If
        
        If OB_SI130.Value = True Then
            ActiveCell.Offset(0, 317) = 1
            ActiveCell.Offset(0, 318) = 0
        ElseIf OB_NO130.Value = True Then
            ActiveCell.Offset(0, 317) = 0
            ActiveCell.Offset(0, 318) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON UNA PUERTA DE EMERGENCIA DE 1.20 MTS. DE ANCHO MÍNIMO HACIA EL EXTERIOR")
        End If
        
        If OB_SI131.Value = True Then
            ActiveCell.Offset(0, 319) = 1
            ActiveCell.Offset(0, 320) = 0
        ElseIf OB_NO131.Value = True Then
            ActiveCell.Offset(0, 319) = 0
            ActiveCell.Offset(0, 320) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON RAMPAS DE 6% DE PENDIENTE MÁXIMA CON PISO ANTIDERRAPANTE")
        End If
        
        If OB_SI132.Value = True Then
            ActiveCell.Offset(0, 321) = 1
            ActiveCell.Offset(0, 322) = 0
        ElseIf OB_NO132.Value = True Then
            ActiveCell.Offset(0, 321) = 0
            ActiveCell.Offset(0, 322) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON RAMPAS DE 6% DE PENDIENTE MÁXIMA CON PISO ANTIDERRAPANTE")
        End If
        
        If OB_SI133.Value = True Then
            ActiveCell.Offset(0, 323) = 1
            ActiveCell.Offset(0, 324) = 0
        ElseIf OB_NO133.Value = True Then
            ActiveCell.Offset(0, 323) = 0
            ActiveCell.Offset(0, 324) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON ILUMINACIÓN EN LA RUTA DE EVACUACIÓN EMERGENTE")
        End If
        
        If OB_SI134.Value = True Then
            ActiveCell.Offset(0, 325) = 1
            ActiveCell.Offset(0, 326) = 0
        ElseIf OB_NO134.Value = True Then
            ActiveCell.Offset(0, 325) = 0
            ActiveCell.Offset(0, 326) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON ILUMINACIÓN EN LA RUTA DE EVACUACIÓN EMERGENTE")
        End If
        
        If OB_SI135.Value = True Then
            ActiveCell.Offset(0, 327) = 1
            ActiveCell.Offset(0, 328) = 0
        ElseIf OB_NO135.Value = True Then
            ActiveCell.Offset(0, 327) = 0
            ActiveCell.Offset(0, 328) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON UNA PUERTA DE EMERGENCIA DE 1.20 MTS. DE ANCHO MÍNIMO HACIA ÁREA INTERIOR")
        End If
        
        If OB_SI136.Value = True Then
            ActiveCell.Offset(0, 329) = 1
            ActiveCell.Offset(0, 330) = 0
        ElseIf OB_NO136.Value = True Then
            ActiveCell.Offset(0, 329) = 0
            ActiveCell.Offset(0, 330) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON UNA PUERTA DE EMERGENCIA DE 1.20 MTS. DE ANCHO MÍNIMO HACIA ÁREA INTERIOR")
        End If
        
        If OB_SI137.Value = True Then
            ActiveCell.Offset(0, 331) = 1
            ActiveCell.Offset(0, 332) = 0
        ElseIf OB_NO137.Value = True Then
            ActiveCell.Offset(0, 331) = 0
            ActiveCell.Offset(0, 332) = 1
        Else
            MsgBox ("SELECCIONA SI CUMPLE O NO CON RAMPAS DE 6% DE PENDIENTE MÁXIMA CON BARANDALES")
        End If
        
        If OB_SI138.Value = True Then
            ActiveCell.Offset(0, 333) = 1
            ActiveCell.Offset(0, 334) = 0
        ElseIf OB_NO138.Value = True Then
            ActiveCell.Offset(0, 333) = 0
            ActiveCell.Offset(0, 334) = 1
        Else
            MsgBox ("SELECCIONA SI SE REQUIERE O NO CON RAMPAS DE 6% DE PENDIENTE MÁXIMA CON BARANDALES")
        End If
        
        ActiveCell.Offset(0, 335) = Observaciones24
        ActiveCell.Offset(0, 336) = Observaciones25
        ActiveCell.Offset(0, 337) = Observaciones26
        
    'Una vez guardada la información, borramos la información
      
    RFI = Empty
    Edificio_Inventario = Empty
    Edificio = Empty
    Ubicacion = Empty
    Actividad = Empty
    Tipo_Propiedad = Empty
    Sector = Empty
    Sup_Terreno = Empty
    Sup_Total = Empty
    Num_Edificaciones = Empty
    Folio = Empty
    Num_Niveles = Empty
    Num_Cuerpos = Empty
    Num_Edif_Sin_Acc = Empty
    Usuario_Cedula = Empty
    Año_Construccion = Empty
    Remodelacion = Empty
    Tipo_Inmueble = Empty
    Fecha_Registro = Empty
    Aplica_SI = Empty
    Aplica_NO = Empty
    Describa_Accesibilidad = Empty
    Total_Personal = Empty
    Poblacion_Servicio = Empty
    Estacionamiento_Publico = Empty
    Estacionamiento_Personal = Empty
    Elevador = Empty
    Escaleras = Empty
    Salida_Emergencia = Empty
    Poblacion_Servicio = Empty
    Personal_Discapacidad = Empty
    Poblacion_A_Diario = Empty
    Cajon_Discapacidad1 = Empty
    Cajon_Discapacidad2 = Empty
    Salva_Escalera = Empty
    Escalera_Emergencia = Empty
    Desnivel_Acceso = Empty
    OB_SI1 = Empty
    OB_SI2 = Empty
    OB_SI3 = Empty
    OB_SI4 = Empty
    OB_SI5 = Empty
    OB_SI6 = Empty
    OB_SI7 = Empty
    OB_SI8 = Empty
    OB_SI9 = Empty
    OB_SI10 = Empty
    OB_SI11 = Empty
    OB_SI12 = Empty
    OB_SI13 = Empty
    OB_SI14 = Empty
    OB_SI15 = Empty
    OB_SI16 = Empty
    OB_SI17 = Empty
    OB_SI18 = Empty
    OB_SI19 = Empty
    OB_SI20 = Empty
    OB_SI21 = Empty
    OB_SI22 = Empty
    OB_SI23 = Empty
    OB_SI24 = Empty
    OB_SI25 = Empty
    OB_SI26 = Empty
    OB_SI27 = Empty
    OB_SI28 = Empty
    OB_SI29 = Empty
    OB_SI30 = Empty
    OB_SI31 = Empty
    OB_SI32 = Empty
    OB_SI33 = Empty
    OB_SI34 = Empty
    OB_SI35 = Empty
    OB_SI36 = Empty
    OB_SI37 = Empty
    OB_SI38 = Empty
    OB_SI39 = Empty
    OB_SI40 = Empty
    OB_SI41 = Empty
    OB_SI42 = Empty
    OB_SI43 = Empty
    OB_SI44 = Empty
    OB_SI45 = Empty
    OB_SI46 = Empty
    OB_SI47 = Empty
    OB_SI48 = Empty
    OB_SI49 = Empty
    OB_SI50 = Empty
    OB_SI51 = Empty
    OB_SI52 = Empty
    OB_SI53 = Empty
    OB_SI54 = Empty
    OB_SI55 = Empty
    OB_SI56 = Empty
    OB_SI57 = Empty
    OB_SI58 = Empty
    OB_SI59 = Empty
    OB_SI60 = Empty
    OB_SI61 = Empty
    OB_SI62 = Empty
    OB_SI63 = Empty
    OB_SI64 = Empty
    OB_SI65 = Empty
    OB_SI66 = Empty
    OB_SI67 = Empty
    OB_SI68 = Empty
    OB_SI69 = Empty
    OB_SI70 = Empty
    OB_SI71 = Empty
    OB_SI72 = Empty
    OB_SI73 = Empty
    OB_SI74 = Empty
    OB_SI75 = Empty
    OB_SI76 = Empty
    OB_SI77 = Empty
    OB_SI78 = Empty
    OB_SI79 = Empty
    OB_SI80 = Empty
    OB_SI81 = Empty
    OB_SI82 = Empty
    OB_SI83 = Empty
    OB_SI84 = Empty
    OB_SI85 = Empty
    OB_SI86 = Empty
    OB_SI87 = Empty
    OB_SI88 = Empty
    OB_SI89 = Empty
    OB_SI90 = Empty
    OB_SI91 = Empty
    OB_SI92 = Empty
    OB_SI93 = Empty
    OB_SI94 = Empty
    OB_SI95 = Empty
    OB_SI96 = Empty
    OB_SI97 = Empty
    OB_SI98 = Empty
    OB_SI99 = Empty
    OB_SI100 = Empty
    OB_SI101 = Empty
    OB_SI102 = Empty
    OB_SI103 = Empty
    OB_SI104 = Empty
    OB_SI105 = Empty
    OB_SI106 = Empty
    OB_SI107 = Empty
    OB_SI108 = Empty
    OB_SI109 = Empty
    OB_SI110 = Empty
    OB_SI111 = Empty
    OB_SI112 = Empty
    OB_SI113 = Empty
    OB_SI114 = Empty
    OB_SI115 = Empty
    OB_SI116 = Empty
    OB_SI117 = Empty
    OB_SI118 = Empty
    OB_SI119 = Empty
    OB_SI120 = Empty
    OB_SI121 = Empty
    OB_SI122 = Empty
    OB_SI123 = Empty
    OB_SI124 = Empty
    OB_SI125 = Empty
    OB_SI126 = Empty
    OB_SI127 = Empty
    OB_SI128 = Empty
    OB_SI129 = Empty
    OB_SI130 = Empty
    OB_SI131 = Empty
    OB_SI132 = Empty
    OB_SI133 = Empty
    OB_SI134 = Empty
    OB_SI135 = Empty
    OB_SI136 = Empty
    OB_SI137 = Empty
    OB_SI138 = Empty
    OB_NO1 = Empty
    OB_NO2 = Empty
    OB_NO3 = Empty
    OB_NO4 = Empty
    OB_NO5 = Empty
    OB_NO6 = Empty
    OB_NO7 = Empty
    OB_NO8 = Empty
    OB_NO9 = Empty
    OB_NO10 = Empty
    OB_NO11 = Empty
    OB_NO12 = Empty
    OB_NO13 = Empty
    OB_NO14 = Empty
    OB_NO15 = Empty
    OB_NO16 = Empty
    OB_NO17 = Empty
    OB_NO18 = Empty
    OB_NO19 = Empty
    OB_NO20 = Empty
    OB_NO21 = Empty
    OB_NO22 = Empty
    OB_NO23 = Empty
    OB_NO24 = Empty
    OB_NO25 = Empty
    OB_NO26 = Empty
    OB_NO27 = Empty
    OB_NO28 = Empty
    OB_NO29 = Empty
    OB_NO30 = Empty
    OB_NO31 = Empty
    OB_NO32 = Empty
    OB_NO33 = Empty
    OB_NO34 = Empty
    OB_NO35 = Empty
    OB_NO36 = Empty
    OB_NO37 = Empty
    OB_NO38 = Empty
    OB_NO39 = Empty
    OB_NO40 = Empty
    OB_NO41 = Empty
    OB_NO42 = Empty
    OB_NO43 = Empty
    OB_NO44 = Empty
    OB_NO45 = Empty
    OB_NO46 = Empty
    OB_NO47 = Empty
    OB_NO48 = Empty
    OB_NO49 = Empty
    OB_NO50 = Empty
    OB_NO51 = Empty
    OB_NO52 = Empty
    OB_NO53 = Empty
    OB_NO54 = Empty
    OB_NO55 = Empty
    OB_NO56 = Empty
    OB_NO57 = Empty
    OB_NO58 = Empty
    OB_NO59 = Empty
    OB_NO60 = Empty
    OB_NO61 = Empty
    OB_NO62 = Empty
    OB_NO63 = Empty
    OB_NO64 = Empty
    OB_NO65 = Empty
    OB_NO66 = Empty
    OB_NO67 = Empty
    OB_NO68 = Empty
    OB_NO69 = Empty
    OB_NO70 = Empty
    OB_NO71 = Empty
    OB_NO72 = Empty
    OB_NO73 = Empty
    OB_NO74 = Empty
    OB_NO75 = Empty
    OB_NO76 = Empty
    OB_NO77 = Empty
    OB_NO78 = Empty
    OB_NO79 = Empty
    OB_NO80 = Empty
    OB_NO81 = Empty
    OB_NO82 = Empty
    OB_NO83 = Empty
    OB_NO84 = Empty
    OB_NO85 = Empty
    OB_NO86 = Empty
    OB_NO87 = Empty
    OB_NO88 = Empty
    OB_NO89 = Empty
    OB_NO90 = Empty
    OB_NO91 = Empty
    OB_NO92 = Empty
    OB_NO93 = Empty
    OB_NO94 = Empty
    OB_NO95 = Empty
    OB_NO96 = Empty
    OB_NO97 = Empty
    OB_NO98 = Empty
    OB_NO99 = Empty
    OB_NO100 = Empty
    OB_NO101 = Empty
    OB_NO102 = Empty
    OB_NO103 = Empty
    OB_NO104 = Empty
    OB_NO105 = Empty
    OB_NO106 = Empty
    OB_NO107 = Empty
    OB_NO108 = Empty
    OB_NO109 = Empty
    OB_NO110 = Empty
    OB_NO111 = Empty
    OB_NO112 = Empty
    OB_NO113 = Empty
    OB_NO114 = Empty
    OB_NO115 = Empty
    OB_NO116 = Empty
    OB_NO117 = Empty
    OB_NO118 = Empty
    OB_NO119 = Empty
    OB_NO120 = Empty
    OB_NO121 = Empty
    OB_NO122 = Empty
    OB_NO123 = Empty
    OB_NO124 = Empty
    OB_NO125 = Empty
    OB_NO126 = Empty
    OB_NO127 = Empty
    OB_NO128 = Empty
    OB_NO129 = Empty
    OB_NO130 = Empty
    OB_NO131 = Empty
    OB_NO132 = Empty
    OB_NO133 = Empty
    OB_NO134 = Empty
    OB_NO135 = Empty
    OB_NO136 = Empty
    OB_NO137 = Empty
    OB_NO138 = Empty
    Observaciones1 = Empty
    Observaciones2 = Empty
    Observaciones3 = Empty
    Observaciones4 = Empty
    Observaciones5 = Empty
    Observaciones6 = Empty
    Observaciones7 = Empty
    Observaciones8 = Empty
    Observaciones9 = Empty
    Observaciones10 = Empty
    Observaciones11 = Empty
    Observaciones12 = Empty
    Observaciones13 = Empty
    Observaciones14 = Empty
    Observaciones15 = Empty
    Observaciones16 = Empty
    Observaciones17 = Empty
    Observaciones18 = Empty
    Observaciones19 = Empty
    Observaciones20 = Empty
    Observaciones21 = Empty
    Observaciones22 = Empty
    Observaciones23 = Empty
    Observaciones24 = Empty
    Observaciones25 = Empty
    Observaciones26 = Empty
End Sub
Private Sub Limpiar_Click()
    RFI = Empty
    Edificio_Inventario = Empty
    Edificio = Empty
    Ubicacion = Empty
    Actividad = Empty
    Tipo_Propiedad = Empty
    Sector = Empty
    Sup_Terreno = Empty
    Sup_Total = Empty
    Num_Edificaciones = Empty
    Folio = Empty
    Num_Niveles = Empty
    Num_Cuerpos = Empty
    Num_Edif_Sin_Acc = Empty
    Usuario_Cedula = Empty
    Año_Construccion = Empty
    Remodelacion = Empty
    Tipo_Inmueble = Empty
    Fecha_Registro = Empty
    Aplica_SI = Empty
    Aplica_NO = Empty
    Describa_Accesibilidad = Empty
    Total_Personal = Empty
    Poblacion_Servicio = Empty
    Estacionamiento_Publico = Empty
    Estacionamiento_Personal = Empty
    Elevador = Empty
    Escaleras = Empty
    Salida_Emergencia = Empty
    Poblacion_Servicio = Empty
    Personal_Discapacidad = Empty
    Poblacion_A_Diario = Empty
    Cajon_Discapacidad1 = Empty
    Cajon_Discapacidad2 = Empty
    Salva_Escalera = Empty
    Escalera_Emergencia = Empty
    Desnivel_Acceso = Empty
    OB_SI1 = Empty
    OB_SI2 = Empty
    OB_SI3 = Empty
    OB_SI4 = Empty
    OB_SI5 = Empty
    OB_SI6 = Empty
    OB_SI7 = Empty
    OB_SI8 = Empty
    OB_SI9 = Empty
    OB_SI10 = Empty
    OB_SI11 = Empty
    OB_SI12 = Empty
    OB_SI13 = Empty
    OB_SI14 = Empty
    OB_SI15 = Empty
    OB_SI16 = Empty
    OB_SI17 = Empty
    OB_SI18 = Empty
    OB_SI19 = Empty
    OB_SI20 = Empty
    OB_SI21 = Empty
    OB_SI22 = Empty
    OB_SI23 = Empty
    OB_SI24 = Empty
    OB_SI25 = Empty
    OB_SI26 = Empty
    OB_SI27 = Empty
    OB_SI28 = Empty
    OB_SI29 = Empty
    OB_SI30 = Empty
    OB_SI31 = Empty
    OB_SI32 = Empty
    OB_SI33 = Empty
    OB_SI34 = Empty
    OB_SI35 = Empty
    OB_SI36 = Empty
    OB_SI37 = Empty
    OB_SI38 = Empty
    OB_SI39 = Empty
    OB_SI40 = Empty
    OB_SI41 = Empty
    OB_SI42 = Empty
    OB_SI43 = Empty
    OB_SI44 = Empty
    OB_SI45 = Empty
    OB_SI46 = Empty
    OB_SI47 = Empty
    OB_SI48 = Empty
    OB_SI49 = Empty
    OB_SI50 = Empty
    OB_SI51 = Empty
    OB_SI52 = Empty
    OB_SI53 = Empty
    OB_SI54 = Empty
    OB_SI55 = Empty
    OB_SI56 = Empty
    OB_SI57 = Empty
    OB_SI58 = Empty
    OB_SI59 = Empty
    OB_SI60 = Empty
    OB_SI61 = Empty
    OB_SI62 = Empty
    OB_SI63 = Empty
    OB_SI64 = Empty
    OB_SI65 = Empty
    OB_SI66 = Empty
    OB_SI67 = Empty
    OB_SI68 = Empty
    OB_SI69 = Empty
    OB_SI70 = Empty
    OB_SI71 = Empty
    OB_SI72 = Empty
    OB_SI73 = Empty
    OB_SI74 = Empty
    OB_SI75 = Empty
    OB_SI76 = Empty
    OB_SI77 = Empty
    OB_SI78 = Empty
    OB_SI79 = Empty
    OB_SI80 = Empty
    OB_SI81 = Empty
    OB_SI82 = Empty
    OB_SI83 = Empty
    OB_SI84 = Empty
    OB_SI85 = Empty
    OB_SI86 = Empty
    OB_SI87 = Empty
    OB_SI88 = Empty
    OB_SI89 = Empty
    OB_SI90 = Empty
    OB_SI91 = Empty
    OB_SI92 = Empty
    OB_SI93 = Empty
    OB_SI94 = Empty
    OB_SI95 = Empty
    OB_SI96 = Empty
    OB_SI97 = Empty
    OB_SI98 = Empty
    OB_SI99 = Empty
    OB_SI100 = Empty
    OB_SI101 = Empty
    OB_SI102 = Empty
    OB_SI103 = Empty
    OB_SI104 = Empty
    OB_SI105 = Empty
    OB_SI106 = Empty
    OB_SI107 = Empty
    OB_SI108 = Empty
    OB_SI109 = Empty
    OB_SI110 = Empty
    OB_SI111 = Empty
    OB_SI112 = Empty
    OB_SI113 = Empty
    OB_SI114 = Empty
    OB_SI115 = Empty
    OB_SI116 = Empty
    OB_SI117 = Empty
    OB_SI118 = Empty
    OB_SI119 = Empty
    OB_SI120 = Empty
    OB_SI121 = Empty
    OB_SI122 = Empty
    OB_SI123 = Empty
    OB_SI124 = Empty
    OB_SI125 = Empty
    OB_SI126 = Empty
    OB_SI127 = Empty
    OB_SI128 = Empty
    OB_SI129 = Empty
    OB_SI130 = Empty
    OB_SI131 = Empty
    OB_SI132 = Empty
    OB_SI133 = Empty
    OB_SI134 = Empty
    OB_SI135 = Empty
    OB_SI136 = Empty
    OB_SI137 = Empty
    OB_SI138 = Empty
    OB_NO1 = Empty
    OB_NO2 = Empty
    OB_NO3 = Empty
    OB_NO4 = Empty
    OB_NO5 = Empty
    OB_NO6 = Empty
    OB_NO7 = Empty
    OB_NO8 = Empty
    OB_NO9 = Empty
    OB_NO10 = Empty
    OB_NO11 = Empty
    OB_NO12 = Empty
    OB_NO13 = Empty
    OB_NO14 = Empty
    OB_NO15 = Empty
    OB_NO16 = Empty
    OB_NO17 = Empty
    OB_NO18 = Empty
    OB_NO19 = Empty
    OB_NO20 = Empty
    OB_NO21 = Empty
    OB_NO22 = Empty
    OB_NO23 = Empty
    OB_NO24 = Empty
    OB_NO25 = Empty
    OB_NO26 = Empty
    OB_NO27 = Empty
    OB_NO28 = Empty
    OB_NO29 = Empty
    OB_NO30 = Empty
    OB_NO31 = Empty
    OB_NO32 = Empty
    OB_NO33 = Empty
    OB_NO34 = Empty
    OB_NO35 = Empty
    OB_NO36 = Empty
    OB_NO37 = Empty
    OB_NO38 = Empty
    OB_NO39 = Empty
    OB_NO40 = Empty
    OB_NO41 = Empty
    OB_NO42 = Empty
    OB_NO43 = Empty
    OB_NO44 = Empty
    OB_NO45 = Empty
    OB_NO46 = Empty
    OB_NO47 = Empty
    OB_NO48 = Empty
    OB_NO49 = Empty
    OB_NO50 = Empty
    OB_NO51 = Empty
    OB_NO52 = Empty
    OB_NO53 = Empty
    OB_NO54 = Empty
    OB_NO55 = Empty
    OB_NO56 = Empty
    OB_NO57 = Empty
    OB_NO58 = Empty
    OB_NO59 = Empty
    OB_NO60 = Empty
    OB_NO61 = Empty
    OB_NO62 = Empty
    OB_NO63 = Empty
    OB_NO64 = Empty
    OB_NO65 = Empty
    OB_NO66 = Empty
    OB_NO67 = Empty
    OB_NO68 = Empty
    OB_NO69 = Empty
    OB_NO70 = Empty
    OB_NO71 = Empty
    OB_NO72 = Empty
    OB_NO73 = Empty
    OB_NO74 = Empty
    OB_NO75 = Empty
    OB_NO76 = Empty
    OB_NO77 = Empty
    OB_NO78 = Empty
    OB_NO79 = Empty
    OB_NO80 = Empty
    OB_NO81 = Empty
    OB_NO82 = Empty
    OB_NO83 = Empty
    OB_NO84 = Empty
    OB_NO85 = Empty
    OB_NO86 = Empty
    OB_NO87 = Empty
    OB_NO88 = Empty
    OB_NO89 = Empty
    OB_NO90 = Empty
    OB_NO91 = Empty
    OB_NO92 = Empty
    OB_NO93 = Empty
    OB_NO94 = Empty
    OB_NO95 = Empty
    OB_NO96 = Empty
    OB_NO97 = Empty
    OB_NO98 = Empty
    OB_NO99 = Empty
    OB_NO100 = Empty
    OB_NO101 = Empty
    OB_NO102 = Empty
    OB_NO103 = Empty
    OB_NO104 = Empty
    OB_NO105 = Empty
    OB_NO106 = Empty
    OB_NO107 = Empty
    OB_NO108 = Empty
    OB_NO109 = Empty
    OB_NO110 = Empty
    OB_NO111 = Empty
    OB_NO112 = Empty
    OB_NO113 = Empty
    OB_NO114 = Empty
    OB_NO115 = Empty
    OB_NO116 = Empty
    OB_NO117 = Empty
    OB_NO118 = Empty
    OB_NO119 = Empty
    OB_NO120 = Empty
    OB_NO121 = Empty
    OB_NO122 = Empty
    OB_NO123 = Empty
    OB_NO124 = Empty
    OB_NO125 = Empty
    OB_NO126 = Empty
    OB_NO127 = Empty
    OB_NO128 = Empty
    OB_NO129 = Empty
    OB_NO130 = Empty
    OB_NO131 = Empty
    OB_NO132 = Empty
    OB_NO133 = Empty
    OB_NO134 = Empty
    OB_NO135 = Empty
    OB_NO136 = Empty
    OB_NO137 = Empty
    OB_NO138 = Empty
    Observaciones1 = Empty
    Observaciones2 = Empty
    Observaciones3 = Empty
    Observaciones4 = Empty
    Observaciones5 = Empty
    Observaciones6 = Empty
    Observaciones7 = Empty
    Observaciones8 = Empty
    Observaciones9 = Empty
    Observaciones10 = Empty
    Observaciones11 = Empty
    Observaciones12 = Empty
    Observaciones13 = Empty
    Observaciones14 = Empty
    Observaciones15 = Empty
    Observaciones16 = Empty
    Observaciones17 = Empty
    Observaciones18 = Empty
    Observaciones19 = Empty
    Observaciones20 = Empty
    Observaciones21 = Empty
    Observaciones22 = Empty
    Observaciones23 = Empty
    Observaciones24 = Empty
    Observaciones25 = Empty
    Observaciones26 = Empty
    
End Sub
Private Sub ScrollBar1_Change()

    ' Desplazar el Frame hacia arriba o abajo
    FrameContenido.Top = -ScrollBar1.Value * 10 ' Ajusta la sensibilidad aquí

End Sub
'Colocar las observaciones en mayúsculas
Private Sub Observaciones1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones10_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones11_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones12_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones13_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones14_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones15_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones16_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones17_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones18_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones19_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones20_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones21_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones22_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones23_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones24_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones25_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Observaciones26_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Total_Personal_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
        Select Case KeyAscii
        Case 48 To 57    ' Dígitos del 0 al 9
        Case 8           ' Backspace
        Case 44, 46      ' Coma o punto decimal
            ' Evitar más de un punto/coma decimal
            If InStr(Total_Personal.Text, Chr(KeyAscii)) > 0 Then
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Ubicacion_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub UserForm_Initialize()
       With Actividad
        .Clear
        .AddItem "Educativo"
        .AddItem "Administrativo"
        .AddItem "Hospital"
        .AddItem "Servicios"
        .AddItem "Otro"
    End With
    
       With Tipo_Propiedad
        .Clear
        .AddItem "Propiedad Federal"
        .AddItem "Arrendado"
        .AddItem "Comodato"
        .AddItem "Prestado"
        .AddItem "En Destino"
        .AddItem "Otro"
    End With
    
       With Sector
        .Clear
        .AddItem "Educativo"
        .AddItem "Público"
        .AddItem "Salud"
        .AddItem "Otro"
    End With
End Sub
