// 1. Obtener las conferencias del JSON
async function obtenerConferencias() {
    const respuesta = await fetch("conferencias.json");

    if(!respuesta.ok) {
        throw new Error("No se pudo cargar conferencias.json")
    }

    return await respuesta.json();
}

// 2. Crear el enlace para Outlook
function crearEnlaceOutlook(conferencia) {
    const parametros = new URLSearchParams({
        subject: conferencia.titulo,

        startdt:
            `${conferencia.fecha}T${conferencia.horaInicio}:00`,

        enddt:
            `${conferencia.fecha}T${conferencia.horaFin}:00`,

        location:
            `${conferencia.campus} - ${conferencia.ubicacion}`,

        body:
            `Ponente: ${conferencia.ponente}\n` +
            `Área: ${conferencia.area}\n\n` +
            conferencia.descripcion
    });

    return (
        "https://outlook.office.com/calendar/0/deeplink/compose?" +
        parametros.toString()
    );
}

// 3. Crear la tarjeta de cada conferencia
function crearTarjeta(conferencia) {
    const tarjeta = document.createElement("article");
    tarjeta.className = "conference-row";

    // Información principal
    const info = document.createElement("div");
    info.className = "conference-info";

    const categoria = document.createElement("span");
    categoria.className = "category-tag technology";
    categoria.textContent = conferencia.area;

    const titulo = document.createElement("h3");
    titulo.textContent = conferencia.titulo;

    const ponente = document.createElement("p");
    ponente.textContent = `Ponente: ${conferencia.ponente}`;

    info.append(categoria, titulo, ponente);

    // Ubicación
    const ubicacion = document.createElement("div");
    ubicacion.className = "location-info";

    const campus = document.createElement("strong");
    campus.textContent = conferencia.campus;

    const lugar = document.createElement("span");
    lugar.textContent = conferencia.ubicacion;

    ubicacion.append(campus, lugar);

    // Fecha y horario
    const horario = document.createElement("div");
    horario.className = "time-info";

    const horas = document.createElement("strong");
    horas.textContent =
        `${conferencia.horaInicio} - ${conferencia.horaFin}`;

    const fecha = document.createElement("span");
    fecha.textContent = conferencia.fecha;

    horario.append(horas, fecha);

    // Botón de Outlook
    const boton = document.createElement("a");
    boton.href = crearEnlaceOutlook(conferencia);
    boton.target = "_blank";
    boton.rel = "noopener noreferrer";
    boton.className = "outlook-button";
    boton.textContent = "Agregar a Outlook";

    // Añadir todos los elementos a la tarjeta
    tarjeta.append(info, ubicacion, horario, boton);

    return tarjeta;
}

// Cargar y mostrar las conferencias
async function mostrarConferencias() {
    const contenedor = document.getElementById(
        "conferencias-container"
    );

    const contador = document.getElementById("results-count");

    try {
        const conferencias = await obtenerConferencias();

        contenedor.replaceChildren();

        conferencias.forEach(conferencia => {
            const tarjeta = crearTarjeta(conferencia);
            contenedor.appendChild(tarjeta);
        });

        contador.textContent =
            `${conferencias.length} conferencias`;

    } catch (error) {
        console.error(error);

        contador.textContent = "No se pudieron cargar las conferencias";

        contenedor.textContent =
            "Ocurrió un error al cargar los eventos. Intenta nuevamente.";
    }
}

// 5. Ejecutar al cargar la página
mostrarConferencias();