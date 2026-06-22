Justificación técnica de las decisiones

Clientes: Los valores nulos en email y ciudad fueron reemplazados por valores controlados ("sin_email@empresa.com" y "Sin datos") para preservar el historial de clientes y evitar pérdida de información relevante para futuros análisis.

Productos: El registro con precio nulo fue eliminado debido a que el precio es un atributo crítico para el cálculo de ingresos y márgenes. El registro con categoría nula fue reemplazado por "Sin Categoría" para mantener el producto disponible en el modelo sin asignarlo incorrectamente a una categoría existente.

Modelo: Se aplicó una estructura de modelo estrella utilizando las dimensiones Dim_Clientes, Dim_Productos y Dim_Categorias, y la tabla de hechos Fact_Ventas. Además, se enriqueció Fact_Ventas mediante un Merge con Dim_Productos para incorporar el nombre y la categoría del producto.

name=Dim_Clientes
let
    // Source: leer la tabla cargada desde el Excel del libro
    Source = Excel.CurrentWorkbook(){[Name="clientes"]}[Content],

    // Renombrar columnas a nomenclatura profesional en español y snake_case
    // Esto facilita el mapeo con el modelo y evita nombres con espacios que rompen fórmulas DAX.
    RenamedCols = Table.RenameColumns(Source, {
        {"ID", "id_cliente"},
        {"Nombre", "nombre_cliente"},
        {"Email", "email"},
        {"Ciudad", "ciudad"},
        {"FechaRegistro", "fecha_registro"}
    }, MissingField.Ignore),

    // Corregir tipos antes de operar sobre valores nulos/duplicados
    TypesFixed = Table.TransformColumnTypes(RenamedCols, {
        {"id_cliente", Int64.Type},
        {"nombre_cliente", type text},
        {"email", type text},
        {"ciudad", type text},
        {"fecha_registro", type date}
    }),

    // Eliminar duplicados por la PK id_cliente para garantizar unicidad en la dimensión.
    // Mantener una sola fila por clave evita problemas en relaciones 1:* en el modelo.
    RemovedDuplicates = Table.Distinct(TypesFixed, {"id_cliente"}),

    // Detectar y marcar nulos: preferimos imputar y documentar en vez de eliminar filas con ventas asociadas.
    // email -> "sin_email@nodisponible.local", crear indicador email_missing
    EmailReplaced = Table.AddColumn(RemovedDuplicates, "email_missing", each (Record.FieldValues(_){List.PositionOf(Table.ColumnNames(RemovedDuplicates),"email")} = null)?, type logical),
    EmailFilled = Table.TransformColumns(Table.TransformColumns(EmailReplaced, {{"email", each if _ = null or Text.Trim(_) = "" then "sin_email@nodisponible.local" else _, type text}}), {}, null),

    // ciudad -> "Sin_Ciudad", crear indicador ciudad_missing
    CiudadReplaced = Table.AddColumn(EmailFilled, "ciudad_missing", each (Record.FieldValues(_){List.PositionOf(Table.ColumnNames(EmailFilled),"ciudad")} = null)?, type logical),
    CiudadFilled = Table.TransformColumns(Table.TransformColumns(CiudadReplaced, {{"ciudad", each if _ = null or Text.Trim(_) = "" then "Sin_Ciudad" else _, type text}}), {}, null),

    // Reordenar columnas y asegurar tipos finales
    FinalTypes = Table.TransformColumnTypes(CiudadFilled, {
        {"id_cliente", Int64.Type},
        {"nombre_cliente", type text},
        {"email", type text},
        {"ciudad", type text},
        {"fecha_registro", type date},
        {"email_missing", type logical},
        {"ciudad_missing", type logical}
    })
in
    FinalTypes
