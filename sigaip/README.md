# SIGAIP – Reservas del Aula de Innovación Pedagógica

Web app en Google Apps Script + Google Sheets, pensada para celular.

## Instalación (5 min)
1. Crea una hoja de cálculo de Google nueva → **Extensiones → Apps Script**.
2. Copia `Code.gs`, `Index.html` (y en *Configuración del proyecto* activa "Mostrar appsscript.json" y pega `appsscript.json`).
3. Ejecuta la función **`setup`** una vez (autoriza permisos). Crea las hojas `Recursos`, `Docentes`, `Bloques` y `Reservas`.
4. Edita las hojas:
   - **Recursos**: ya trae 35 laptops, 7 proyectores, 4 parlantes, 1 sala de cómputo.
   - **Docentes**: pega tu lista (Docente, Área curricular, Activo SI/NO).
   - **Bloques**: pon las horas reales de tu horario (una fila que contenga "Recreo" se bloquea sola).
5. **Implementar → Nueva implementación → Aplicación web** (Ejecutar como: yo · Acceso: cualquiera). Comparte el enlace por WhatsApp.

## Qué hace
- Reserva por **día** (esta semana + la próxima), **recurso** y **una o varias horas**.
- Laptops/proyectores/parlantes se piden por **cantidad**; el sistema descuenta lo ya reservado y no permite sobrepasar el total.
- La sala de cómputo es de 1 unidad: si está ocupada, la hora aparece bloqueada.
- Evita reservas dobles simultáneas (`LockService`).
- Cada docente ve y **cancela** sus reservas en "Mis reservas". Recuerda su nombre en el celular.
- Todo queda en la hoja `Reservas`, donde puedes filtrar por día o docente.

## Pendiente de decidir
- Autenticación (hoy cualquiera con el enlace puede elegir cualquier nombre; se puede añadir un PIN por docente).
- Aviso por correo/WhatsApp al reservar.
