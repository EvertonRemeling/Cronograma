document.addEventListener('DOMContentLoaded', () => {
    const scheduleTableContainer = document.getElementById('scheduleTableContainer');
    const storedCalendar = localStorage.getItem('scheduledVisitsData');

    if (!storedCalendar) {
        scheduleTableContainer.innerHTML = '<p class="message">Nenhuma agenda encontrada. Por favor, crie e salve um cronograma na página principal.</p>';
        return;
    }

    let calendarAllocations = JSON.parse(storedCalendar);

    let allScheduledActivities = [];
    for (const dateString in calendarAllocations) {
        calendarAllocations[dateString].forEach(activity => {
            allScheduledActivities.push({
                id: activity.id,
                description: activity.description,
                module: activity.module || 'N/A',
                date: activity.date,
                shift: activity.shift || 'N/A',
                technicians: activity.technicians || []
            });
        });
    }

    if (allScheduledActivities.length === 0) {
        scheduleTableContainer.innerHTML = '<p class="message">Nenhuma atividade agendada no cronograma.</p>';
        return;
    }

    allScheduledActivities.sort((a, b) => {
        const dateA = new Date(a.date);
        const dateB = new Date(b.date);
        if (dateA - dateB !== 0) {
            return dateA - dateB;
        }
        const shiftOrder = ['Manhã (08h-12h)', 'Tarde (13h-17h)'].map(s => s.split(' ')[0]);
        const shiftA = a.shift ? a.shift.split(' ')[0] : '';
        const shiftB = b.shift ? b.shift.split(' ')[0] : '';
        return shiftOrder.indexOf(shiftA) - shiftOrder.indexOf(shiftB);
    });

    let tableHtml = `
        <table>
            <thead>
                <tr>
                    <th>Data</th>
                    <th>Turno</th>
                    <th>Módulo</th>
                    <th>Assunto</th>
                    <th>Técnicos Envolvidos</th>
                </tr>
            </thead>
            <tbody>
    `;

    allScheduledActivities.forEach(activity => {
        tableHtml += `
            <tr>
                <td>${new Date(activity.date).toLocaleDateString('pt-BR')}</td>
                <td>${activity.shift.split(' ')[0]}</td>
                <td>${activity.module}</td>
                <td>${activity.description}</td>
                <td>${activity.technicians && activity.technicians.length > 0 ? activity.technicians.join(', ') : 'N/A'}</td>
            </tr>
        `;
    });

    tableHtml += `
            </tbody>
        </table>
    `;

    scheduleTableContainer.innerHTML = tableHtml;

    // NOVIDADE: Verifica se a página foi aberta com a intenção de imprimir
    const printMode = sessionStorage.getItem('printMode');
    if (printMode === 'true') {
        sessionStorage.removeItem('printMode'); // Limpa a flag para impressões futuras
        window.print(); // Dispara a impressão
    }
});