document.addEventListener('DOMContentLoaded', () => {
    const boardsContainer = document.getElementById('boardsContainer');
    const calendarGrid = document.getElementById('calendarGrid');
    const calendarMonthYear = document.getElementById('calendarMonthYear');
    const prevMonthBtn = document.getElementById('prevMonthBtn');
    const nextMonthBtn = document.getElementById('nextMonthBtn');
    const printScheduleBtn = document.getElementById('printScheduleBtn');
    const integrateSiclaBtn = document.getElementById('integrateSiclaBtn');
    const boardDragGhost = document.getElementById('boardDragGhost');
    const moduleFilter = document.getElementById('moduleFilter');
    const technicianSelector = document.getElementById('technicianSelector');

    // --- Dados do Anexo (resultado da análise da planilha) ---
    const rawSpreadsheetResult = `{"comment":"The request is to retrieve all data from the table 'planilha1'.","data":[{"number_of_columns":8,"columns":["modulo","adicional","tipo","integracoes","item_de_go_live","menu","item","sequencia"],"number_of_records":98,"records":[["CTB","CTB","Configuração","","","1.2.M","Manutenção Consolidade de Empresas",1],["CTB","CTB","Configuração","","","1.2.M","Manutenção Consolidade de Empresas - CTB - Dados Adicionais",1],["CTB","CTB","Configuração","","","1.2.M","Manutenção Consolidade de Empresas -  CTB - Plano de Contas",1],["CTB","CTB","Configuração","","","1.2.M","Manutenção Consolidade de Empresas -  CTB - Lançamentos",1],["CTB","CTB","Treinamento","","","1.3.M","TAB 67-  Totalizadores Demonstrações Contábeis",1],["CTB","CTB","Treinamento","","","1.3.M","TAB 82-  Incentivos Fiscais/Deduções",1],["CTB","CTB","Treinamento","","","1.3.M","TAB 83-  Totalizadores Apuração Csll/Irpj",1],["CTB","CTB","Treinamento","","","1.3.M","TAB 84-  Configuração Apuração Csll/Irpj",1],["CTB","CTB","Treinamento","","","1.3.M","TAB 87-  Totalizadores Lucro da Exploração",1],["CTB","CCC","Treinamento","","","1.4.M","TAB 02-  Centro de Custo",1],["CTB","CTB","Treinamento","","","1.4.M","TAB 87-  Filiais da Contabilidade",1],["CTB","CCC","Treinamento","","","1.4.M","TAB 87-  Rateio Padrão dos Centros de Custos",1],["CTB","CTB","Treinamento","","","1.6-H","Manutenção Consolidada de Contas Contábeis",2],["CTB","CTB","Treinamento","","","1.6-L","Listagem de Contas Contábeis",2],["CTB","CTB","Treinamento","","","1.6-M","Importação de Contas Contábeis",2],["CTB","CTB","Treinamento","","","1.6-X","Cópia de Configuração do Plano de Contas",2],["CTB","CTB","Treinamento","","","1.6-V","Vinculação de Tabela SPED de Contas Contábeis",2],["CTB","CTB","Treinamento","","","1.6-T","Configuração da Totalizadores do Dem./Apuração",2],["CTB","CTB","Treinamento","","","1.6-U","Vinculação de Filial à Contas Contábeis",3],["CTB","CTB","Treinamento","","","1.8.I","Inclusão de Históricos",3],["CTB","CTB","Treinamento","","","1.8.A","Alteração de Históricos",3],["CTB","CTB","Treinamento","","","1.9-I","Inclusão de Lançamentos Padrões",3],["CTB","CTB","Treinamento","","","1.9-A","Alteração de Lançamentos Padrões",3],["CTB","CTB","Treinamento","","","1.9-L","Listagem de Lançamentos Padrões",3],["CTB","CTB","Treinamento","","","2.1.M","Manutenção Consolidada de Lançamentos Contábeis",3],["CTB","CTB","Treinamento","","","2.1.I","Inclusão  Lançamentos Contábeis",4],["CTB","CTB","Treinamento","","","2.1.L","Listagem de Lançamentos Contábeis",4],["CTB","CTB","Treinamento","","","2.1.S","Implantação de Saldos para Balanço",4],["CTB","CTB","Treinamento","","","2.1.G","Geração de Contabilidade em Outra Moeda",4],["CTB","CTB","Treinamento","","","2.2","Gerção de Contabilidade em Outras Moedas",4],["CTB","CTB","Treinamento","","","2.3.M","Manutenção Consolidada de Movimentos de Caixa/CC",4],["CTB","CTB","Treinamento","","","2.4-B","Bloqueia/Desbloqueia Lotes",5],["CTB","CTB","Treinamento","","","2.4-J","Junçõa de Lotes",5],["CTB","CTB","Treinamento","","","2.4-M","Mudanças do Mês do Lote",5],["CTB","CTB","Treinamento","","","2.4-T","Transferência de Lançamentos de Lotes",5],["CTB","CTB","Treinamento","","","2.4-U","Transferência de Lançamentos de Contas",5],["CTB","CTB","Treinamento","","","2.4-N","Mudança do Nùmero do Lote",5],["CTB","CTB","Treinamento","","","2.4-V","Verificação de Lotes/Contas",6],["CTB","CCC","Treinamento","","","2.4-R","Verificação dos Saldos do Centro de Custo",6],["CTB","CTB","Treinamento","","","2.4-X","Exclusão de Lotes",6],["CTB","CTB","Treinamento","","","2.4-A","Cópia de Lotes",6],["CTB","CTB","Treinamento","","","2.5.I","Importaçao/Digitação Extrato",6],["CTB","CTB","Treinamento","","","2.5.L","Listagem Conciliação",6],["CTB","CTB","Treinamento","","","2.5.C","Conciliação Bancária",7],["CTB","CTB","Treinamento","","","2.5M","Configuração do Extrato Bancário",7],["CTB","CTB","Treinamento","","","2.6.M","Manutenção Consolidade de Despesas Antecipadas",7],["CTB","CTB","Treinamento","","","2.6.I","Integração Mensal de Apropriações",7],["CTB","CTB","Treinamento","","","3.1.A","Atualização de Contas Sintéticas de Balancete",7],["CTB","CTB","Treinamento","","","3.1.C","Balancete Compartivo",7],["CTB","CTB","Treinamento","","","3.1.L","Listagem do Balancete",8],["CTB","CTB","Treinamento","","","3.1.D","Listagem do Balancete de Diversos Períodos",8],["CTB","CTB","Treinamento","","","3.2.L","Listagem do Razão Oficial",8],["CTB","CTB","Treinamento","","","3.3","Análise Vertical",8],["CTB","CTB","Treinamento","","","3.4","Análise a Adicionar/Excluir do Resultado Fiscal",8],["CTB","CTB","Treinamento","","","3.5","Conciliação Contábil",8],["CTB","CTB","Treinamento","","","3.6","Auditoria Contábil",8],["CTB","CTB","Treinamento","","","3.7","Demonstrações Contábeis",9],["CTB","CTB","Treinamento","","","4.1.L","Listagem do Diário",9],["CTB","CTB","Treinamento","","","4.1.T","Emissão do Livro Diário Completo",9],["CTB","CTB","Treinamento","","","4.1.S","Geraçõa do Arquivo Demonst. Complementares SPED",9],["CTB","CTB","Treinamento","","","4.2.A","Atualização de Contas Sintéticas de Balanços",9],["CTB","CTB","Treinamento","","","4.2.C","Balanço para Banco",9],["CTB","CTB","Treinamento","","","4.2.L","Listagem do Balanço Diário",9],["CTB","CTB","Treinamento","","","4.3.G","Geração da DRE",1],["CTB","CTB","Treinamento","","","4.3.L","Listagem da DRE - Diário",10],["CTB","CTB","Treinamento","","","4.3.C","Listagem da DRE - Comparativa",10],["CTB","CTB","Treinamento","","","4.3.P","Listagem da DRE de Diversos Periodos",10],["CTB","CTB","Treinamento","","","4.3.Q","Gráfico da DRE de Diversos Períodos",10],["CTB","CTB","Treinamento","","","4.3.M","Listagem DMPL",10],["CTB","CTB","Treinamento","","","4.3.U","Listagem DLPA",10],["CTB","CTB","Treinamento","","","4.3.X","Listagem DFC",10],["CTB","CTB","Treinamento","","","4.3.V","Listagem DVA",11],["CTB","CTB","Treinamento","","","4.3.N","Listagem DRA",11],["CTB","CTB","Treinamento","","","4.3.E","Listagem EBIT , EBITDA, LAJIDA",11],["CTB","CTB","Treinamento","","","4.4","Livro Caixa De Atividade Rural",11],["CTB","CTB","Treinamento","","","4.5.E","Encerrameto do Exercicio",11],["CTB","CTB","Treinamento","","","4.6.G","Geração de Exercicio Seguinte",11],["CTB","CTB","Treinamento","","","4.6.A","Atulizaçaõ do Balanço de Abertura",11],["CTB","CTB","Treinamento","","","4.7","Demonstrativo de Sobras/Perdas",11],["CTB","CTB","Treinamento","","","5.1.S","Geração de Arquivos SPED - ECD",11],["CTB","CTB","Treinamento","","","5.1.D","Geração de Arquivos SPED - ECF",11],["CTB","CTB","Treinamento","","","5.2.X","Importação de Plano de Contas Via Arquivo XLS",11],["CTB","CTB","Treinamento","","","5.2.I","Importação de Lançamentos Via Arquivo XLS",11],["CTB","CTB","Treinamento","","","5.2.S","Importação de Movimentos de Arquivo SPED - ECD",11],["CTB","CTB","Treinamento","","","5.3","Geração de Plano Consolidado",11],["CTB","CTB","Treinamento","","","5.4","Transferência de Código de Cadastros",11],["CTB","CTB","Treinamento","","","5.5","Geraçõa de Silca com Situação Especial",12],["CTB","CTB","Treinamento","","","5.6","Geração de Planilha",12],["CTB","CTB","Treinamento","","","5.8","Geração de Arquivos para TCE/RS",12],["CTB","CCC","Treinamento","","","6.1.L","Listagem do Mapa de Custos",12],["CTB","CCC","Treinamento","","","6.1.M","Diário Auxiliar de Despesas",12],["CTB","CCC","Treinamento","","","6.1.C","Mapa Comparativo",12],["CTB","CCC","Treinamento","","","6.2.L","Listagem Prévia de Custos",12],["CTB","CCC","Treinamento","","","6.2.P","Prévia de Custos de Receitas /Despesas",13],["CTB","CCC","Treinamento","","","3.3.A","Alteração Geral de Centros de Custos Contábeis",13],["CTB","LALUR","Treinamento","","","7.1.","CSLL/IRPJ",13],["CTB","CTB","Treinamento","","","7.3","Apuração de Lucro Estimado/Presumido",13],["CTB","CTB","Treinamento","","","7.3","Apuração de Valores Transfere Princing",13]]}],"total_number_of_tabs":1}`;

    let allAvailableBoards = [];
    let currentBoardsToDisplay = [];

    let displayDate = new Date();
    let displayMonth = displayDate.getMonth();
    let displayYear = displayDate.getFullYear();

    let calendarAllocations = {};

    const availableTechnicians = ['João Silva', 'Maria Souza', 'Carlos Oliveira', 'Ana Costa', 'Pedro Almeida'];
    const shifts = ['Manhã (08h-12h)', 'Tarde (13h-17h)'];

    function parseRawSpreadsheetData() {
        const parsedData = JSON.parse(rawSpreadsheetResult);
        const records = parsedData.data[0].records;

        let groupedByModuleAndSequence = {};
        let uniqueModules = new Set();
        let initialBoards = [];

        records.forEach(record => {
            const modulo = record[0];
            const sequencia = record[7];
            const item = record[6];

            uniqueModules.add(modulo);

            if (!groupedByModuleAndSequence[modulo]) {
                groupedByModuleAndSequence[modulo] = {};
            }
            if (!groupedByModuleAndSequence[modulo][sequencia]) {
                const boardId = `${modulo}-seq${sequencia}`;
                groupedByModuleAndSequence[modulo][sequencia] = {
                    id: boardId,
                    title: `Visita ${sequencia} (${modulo})`,
                    activities: []
                };
            }
            groupedByModuleAndSequence[modulo][sequencia].activities.push({
                id: `card-${modulo}-${sequencia}-${groupedByModuleAndSequence[modulo][sequencia].activities.length + 1}`,
                description: item,
                module: modulo
            });
        });

        const finBoards = [
            {
                id: 'FIN-seq1',
                title: 'Visita 1 (FIN)',
                activities: [
                    { id: 'card-FIN-1-1', description: 'Reunião de Alinhamento Financeiro', module: 'FIN' },
                    { id: 'card-FIN-1-2', description: 'Coleta de Dados de Fluxo de Caixa', module: 'FIN' }
                ]
            },
            {
                id: 'FIN-seq2',
                title: 'Visita 2 (FIN)',
                activities: [
                    { id: 'card-FIN-2-1', description: 'Análise de Balanço Patrimonial', module: 'FIN' },
                    { id: 'card-FIN-2-2', description: 'Discussão de KPIs Financeiros', module: 'FIN' }
                ]
            }
        ];
        finBoards.forEach(board => {
            const modulo = board.id.split('-')[0];
            const sequencia = board.id.split('-')[1].replace('seq', '');
            if (!groupedByModuleAndSequence[modulo]) groupedByModuleAndSequence[modulo] = {};
            groupedByModuleAndSequence[modulo][sequencia] = board;
            uniqueModules.add(modulo);
        });


        for (const moduloKey in groupedByModuleAndSequence) {
            for (const sequenciaKey in groupedByModuleAndSequence[moduloKey]) {
                initialBoards.push(groupedByModuleAndSequence[moduloKey][sequenciaKey]);
            }
        }

        initialBoards.sort((a, b) => {
            const moduloA = a.id.split('-')[0];
            const moduloB = b.id.split('-')[0];
            if (moduloA !== moduloB) {
                return moduloA.localeCompare(moduloB);
            }
            const seqA = parseInt(a.id.split('-')[1].replace('seq', ''));
            const seqB = parseInt(b.id.split('-')[1].replace('seq', ''));
            return seqA - seqB;
        });

        return { initialBoards, uniqueModules: Array.from(uniqueModules).sort() };
    }

    function createCardElement(activity) {
        const card = document.createElement('div');
        card.classList.add('card');
        card.setAttribute('data-id', activity.id);
        card.setAttribute('data-description', activity.description);
        if (activity.module) card.setAttribute('data-module', activity.module);
        if (activity.technicians) card.setAttribute('data-technicians', activity.technicians.join(','));
        if (activity.shift) card.setAttribute('data-shift', activity.shift);

        let cardContent = activity.description;
        if (activity.technicians && activity.technicians.length > 0) {
            cardContent += ` <br><small>(${activity.technicians.join(', ')})</small>`;
        }
        card.innerHTML = cardContent;
        return card;
    }

    function renderBoards() {
        boardsContainer.innerHTML = '';
        currentBoardsToDisplay.forEach(board => {
            const boardColumn = document.createElement('div');
            boardColumn.classList.add('board-column');
            boardColumn.id = `board-column-${board.id}`;

            const boardTitle = document.createElement('h3');
            boardTitle.textContent = board.title;
            boardTitle.draggable = true;
            boardTitle.setAttribute('data-board-id', board.id);
            boardColumn.appendChild(boardTitle);

            const boardItems = document.createElement('div');
            boardItems.classList.add('board-items');
            boardItems.id = `board-items-${board.id}`;
            boardColumn.appendChild(boardItems);

            board.activities.forEach(activity => {
                const card = createCardElement(activity);
                boardItems.appendChild(card);
            });

            boardsContainer.appendChild(boardColumn);
        });
        initSortableForBoards();
        addBoardDragListeners();
    }

    function renderCalendar() {
        calendarMonthYear.textContent = new Date(displayYear, displayMonth).toLocaleString('pt-BR', { month: 'long', year: 'numeric' });
        
        const dayElements = calendarGrid.querySelectorAll('.calendar-day');
        dayElements.forEach(day => day.remove());

        const firstDayOfMonth = new Date(displayYear, displayMonth, 1).getDay();
        const daysInMonth = new Date(displayYear, displayMonth + 1, 0).getDate();

        for (let i = 0; i < firstDayOfMonth; i++) {
            const emptyDay = document.createElement('div');
            emptyDay.classList.add('calendar-day', 'empty');
            calendarGrid.appendChild(emptyDay);
        }

        for (let dayNum = 1; dayNum <= daysInMonth; dayNum++) {
            const date = new Date(displayYear, displayMonth, dayNum);
            const dateString = date.toISOString().split('T')[0];
            
            const calendarDay = document.createElement('div');
            calendarDay.classList.add('calendar-day');
            calendarDay.setAttribute('data-date', dateString);
            calendarDay.id = `day-${dateString}`;

            const dayNumberSpan = document.createElement('span');
            dayNumberSpan.classList.add('day-number');
            dayNumberSpan.textContent = dayNum;
            calendarDay.appendChild(dayNumberSpan);

            shifts.forEach(shiftName => {
                const shiftSlot = document.createElement('div');
                shiftSlot.classList.add('shift-slot');
                shiftSlot.setAttribute('data-shift', shiftName);
                const shiftTitle = document.createElement('h4');
                shiftTitle.textContent = shiftName.split(' ')[0];
                shiftSlot.appendChild(shiftTitle);
                calendarDay.appendChild(shiftSlot);
            });
            
            if (calendarAllocations[dateString]) {
                const activitiesForThisMonth = calendarAllocations[dateString].filter(activity => {
                    const activityDate = new Date(activity.date);
                    return activityDate.getMonth() === displayMonth && activityDate.getFullYear() === displayYear;
                });

                activitiesForThisMonth.forEach(activity => {
                    const targetShiftSlot = calendarDay.querySelector(`[data-shift="${activity.shift}"]`);
                    if (targetShiftSlot) {
                        targetShiftSlot.appendChild(createCardElement(activity));
                    }
                });
            }

            calendarGrid.appendChild(calendarDay);
        }

        const allDayCells = calendarGrid.querySelectorAll('.calendar-day');
        const numCellsAdded = allDayCells.length;
        const remainingCellsInLastRow = 7 - (numCellsAdded % 7 || 7);
        if (remainingCellsInLastRow < 7 && remainingCellsInLastRow > 0) {
            for (let i = 0; i < remainingCellsInLastRow; i++) {
                const emptyDay = document.createElement('div');
                emptyDay.classList.add('calendar-day', 'empty');
                calendarGrid.appendChild(emptyDay);
            }
        }
        
        const finalCalendarDays = calendarGrid.querySelectorAll('.calendar-day');
        const lastRowStartIndex = finalCalendarDays.length - 7;
        for (let i = lastRowStartIndex; i < finalCalendarDays.length; i++) {
            if (i >= 0) finalCalendarDays[i].classList.add('last-row');
        }

        initSortableForCalendarSlots();
        addCalendarDayDragAndDropListeners();
    }

    function navigateMonth(direction) {
        displayDate.setMonth(displayDate.getMonth() + direction);
        displayMonth = displayDate.getMonth();
        displayYear = displayDate.getFullYear();
        renderCalendar();
    }

    function initSortableForBoards() {
        document.querySelectorAll('.board-items').forEach(boardItems => {
            new Sortable(boardItems, {
                group: {
                    name: 'activities',
                    pull: 'clone',
                    put: false
                },
                animation: 150,
                ghostClass: 'sortable-ghost',
                dragClass: 'sortable-drag',
                onMove: function (evt) {
                    return evt.from !== evt.to;
                },
                onClone: function(evt) {
                    const originalCard = evt.item;
                    const clonedCard = evt.clone;
                    const originalBoardItems = originalCard.closest('.board-items');
                    const boardId = originalBoardItems.id.replace('board-items-', '');
                    const module = boardId.split('-')[0]; 

                    clonedCard.setAttribute('data-original-id', originalCard.getAttribute('data-id'));
                    clonedCard.setAttribute('data-module', module);
                }
            });
        });
    }

    function initSortableForCalendarSlots() {
        shifts.forEach(shiftName => {
            document.querySelectorAll(`.shift-slot[data-shift="${shiftName}"]`).forEach(shiftSlot => {
                new Sortable(shiftSlot, {
                    group: 'activities',
                    animation: 150,
                    ghostClass: 'sortable-ghost',
                    dragClass: 'sortable-drag',
                    onEnd: function (evt) {
                        const draggedCard = evt.item;
                        const targetShiftSlot = evt.to;
                        const targetDateElement = targetShiftSlot.closest('.calendar-day');
                        
                        if (!targetDateElement) { 
                             updateCalendarDataFromDOM();
                             updateBoardDataFromDOM();
                            return;
                        }

                        const newDate = targetDateElement.getAttribute('data-date');
                        const newShift = targetShiftSlot.getAttribute('data-shift');
                        
                        const selectedTechnicians = getSelectedTechnicians();
                        const module = draggedCard.getAttribute('data-module');

                        let cardId = draggedCard.getAttribute('data-id');
                        let originalId = draggedCard.getAttribute('data-original-id');

                        const oldDateElement = evt.from.closest('.calendar-day');
                        if (oldDateElement) {
                            const oldDate = oldDateElement.getAttribute('data-date');
                            if (calendarAllocations[oldDate]) {
                                calendarAllocations[oldDate] = calendarAllocations[oldDate].filter(activity => activity.id !== cardId);
                                if (calendarAllocations[oldDate].length === 0) delete calendarAllocations[oldDate];
                            }
                        }
                        
                        if (originalId && !cardId.startsWith('scheduled-')) {
                            cardId = `scheduled-${originalId}-${Date.now()}`;
                            draggedCard.setAttribute('data-id', cardId);
                        } else if (!cardId || cardId === originalId) {
                            cardId = `scheduled-card-${Date.now()}`;
                            draggedCard.setAttribute('data-id', cardId);
                        }


                        const newActivity = {
                            id: cardId,
                            description: draggedCard.getAttribute('data-description'),
                            module: module,
                            date: newDate,
                            shift: newShift,
                            technicians: selectedTechnicians
                        };
                        
                        if (!calendarAllocations[newDate]) {
                            calendarAllocations[newDate] = [];
                        }
                        calendarAllocations[newDate].push(newActivity);
                        
                        draggedCard.innerHTML = `${newActivity.description} <br><small>(${newActivity.technicians.join(', ')})</small>`;
                        draggedCard.setAttribute('data-technicians', newActivity.technicians.join(','));
                        draggedCard.setAttribute('data-shift', newActivity.shift);
                        draggedCard.setAttribute('data-module', newActivity.module);

                        updateCalendarDataFromDOM();
                        updateBoardDataFromDOM();
                    }
                });
            });
        });
    }

    function updateBoardDataFromDOM() {
        allAvailableBoards.forEach(board => {
            const boardItemsElement = document.getElementById(`board-items-${board.id}`);
            if (boardItemsElement) {
                board.activities = [];
                Array.from(boardItemsElement.children).forEach(cardElement => {
                    board.activities.push({
                        id: cardElement.getAttribute('data-id'),
                        description: cardElement.getAttribute('data-description'),
                        module: cardElement.getAttribute('data-module')
                    });
                });
            }
        });
        saveScheduleToLocalStorage();
    }

    function updateCalendarDataFromDOM() {
        calendarAllocations = {};

        document.querySelectorAll('.calendar-day:not(.empty)').forEach(calendarDayElement => {
            const dateString = calendarDayElement.getAttribute('data-date');
            if (dateString) { 
                calendarDayElement.querySelectorAll('.shift-slot').forEach(shiftSlot => {
                    const shiftName = shiftSlot.getAttribute('data-shift');
                    Array.from(shiftSlot.children).forEach(child => {
                        if (child.classList.contains('card')) {
                            if (!calendarAllocations[dateString]) {
                                calendarAllocations[dateString] = [];
                            }
                            const techs = child.getAttribute('data-technicians');
                            calendarAllocations[dateString].push({
                                id: child.getAttribute('data-id'),
                                description: child.getAttribute('data-description'),
                                module: child.getAttribute('data-module'),
                                date: dateString,
                                shift: shiftName,
                                technicians: techs ? techs.split(',') : []
                            });
                        }
                    });
                });
            }
        });
        saveScheduleToLocalStorage();
    }

    function createBoardDragImage(numActivities) {
        boardDragGhost.textContent = `Arrastando ${numActivities} Visita${numActivities !== 1 ? 's' : ''}`;
        boardDragGhost.style.display = 'block';
        return boardDragGhost;
    }

    function handleBoardDragStart(event) {
        const boardId = event.target.getAttribute('data-board-id');
        const sourceBoard = allAvailableBoards.find(b => b.id === boardId);

        if (sourceBoard && sourceBoard.activities.length > 0) {
            const activitiesToTransfer = sourceBoard.activities.map(activity => ({...activity}));
            
            const transferData = {
                type: 'board-drag',
                boardId: boardId,
                activities: activitiesToTransfer
            };
            event.dataTransfer.setData('application/json', JSON.stringify(transferData));
            event.dataTransfer.effectAllowed = 'move';
            
            const dragImage = createBoardDragImage(activitiesToTransfer.length);
            event.dataTransfer.setDragImage(dragImage, 0, 0);

            event.target.addEventListener('dragend', () => {
                boardDragGhost.style.display = 'none';
            }, { once: true });

        } else {
            event.preventDefault(); 
        }
    }

    function handleCalendarDayDragOver(event) {
        const isBoardDrag = event.dataTransfer.types.includes('application/json');
        
        if (isBoardDrag && event.target.classList.contains('shift-slot')) {
            event.preventDefault(); 
            event.dataTransfer.dropEffect = 'move';
            event.target.classList.add('drag-over');
        } else if (isBoardDrag && event.target.classList.contains('calendar-day')) {
            event.preventDefault();
            event.dataTransfer.dropEffect = 'move';
            event.target.classList.add('drag-over');
        }
    }

    function handleCalendarDayDragLeave(event) {
        if (event.target.classList.contains('shift-slot') || event.target.classList.contains('calendar-day')) {
            event.target.classList.remove('drag-over');
        }
    }

    function handleCalendarDayDrop(event) {
        if (event.target.classList.contains('shift-slot') || event.target.classList.contains('calendar-day')) {
            event.target.classList.remove('drag-over');
        }
        
        const transferData = event.dataTransfer.getData('application/json');
        
        try {
            const data = JSON.parse(transferData);
            if (data && data.type === 'board-drag') {
                let targetShiftSlot = event.target.closest('.shift-slot');
                if (!targetShiftSlot && event.target.classList.contains('calendar-day')) {
                    targetShiftSlot = event.target.querySelector('.shift-slot');
                }
                
                if (!targetShiftSlot) {
                    return;
                }

                event.preventDefault(); 
                event.stopPropagation(); 

                const sourceBoardId = data.boardId;
                const activitiesToMove = data.activities;
                const targetDateString = targetShiftSlot.closest('.calendar-day').getAttribute('data-date');
                const targetShift = targetShiftSlot.getAttribute('data-shift');
                const selectedTechnicians = getSelectedTechnicians();

                const sourceBoardInAll = allAvailableBoards.find(b => b.id === sourceBoardId);
                if (sourceBoardInAll) {
                    sourceBoardInAll.activities = [];
                }

                if (!calendarAllocations[targetDateString]) {
                    calendarAllocations[targetDateString] = [];
                }
                activitiesToMove.forEach(activity => {
                    calendarAllocations[targetDateString].push({
                        id: `scheduled-${activity.id}-${Date.now()}`,
                        description: activity.description,
                        module: activity.module,
                        date: targetDateString,
                        shift: targetShift,
                        technicians: selectedTechnicians
                    });
                });

                filterAndRenderBoards(moduleFilter.value);
                renderCalendar();
                saveScheduleToLocalStorage();
            }
        } catch (e) {
        }
    }

    function addBoardDragListeners() {
        document.querySelectorAll('.board-column h3').forEach(boardTitle => {
            boardTitle.addEventListener('dragstart', handleBoardDragStart);
        });
    }

    function addCalendarDayDragAndDropListeners() {
        document.querySelectorAll('.shift-slot').forEach(shiftSlot => {
            shiftSlot.addEventListener('dragover', handleCalendarDayDragOver);
            shiftSlot.addEventListener('dragleave', handleCalendarDayDragLeave);
            shiftSlot.addEventListener('drop', handleCalendarDayDrop);
        });
        document.querySelectorAll('.calendar-day:not(.empty)').forEach(calendarDay => {
            calendarDay.addEventListener('dragover', handleCalendarDayDragOver);
            calendarDay.addEventListener('dragleave', handleCalendarDayDragLeave);
        });
    }

    function populateModuleFilter(uniqueModules) {
        moduleFilter.innerHTML = '<option value="all">Todos</option>';
        uniqueModules.forEach(moduleName => {
            const option = document.createElement('option');
            option.value = moduleName;
            option.textContent = moduleName;
            moduleFilter.appendChild(option);
        });
    }

    function populateTechnicianSelector() {
        technicianSelector.innerHTML = '';
        availableTechnicians.forEach(tech => {
            const option = document.createElement('option');
            option.value = tech;
            option.textContent = tech;
            technicianSelector.appendChild(option);
        });
    }

    function getSelectedTechnicians() {
        return Array.from(technicianSelector.selectedOptions).map(option => option.value);
    }

    function filterAndRenderBoards(filterValue = 'all') {
        if (filterValue === 'all') {
            currentBoardsToDisplay = allAvailableBoards.filter(board => board.activities.length > 0);
        } else {
            currentBoardsToDisplay = allAvailableBoards.filter(board => 
                board.id.startsWith(`${filterValue}-`) && board.activities.length > 0
            );
        }
        renderBoards();
    }

    function saveScheduleToLocalStorage() {
        localStorage.setItem('scheduledVisitsData', JSON.stringify(calendarAllocations));
        localStorage.setItem('allAvailableBoardsData', JSON.stringify(allAvailableBoards));
        console.log('Dados salvos no localStorage.');
    }

    function loadScheduleFromLocalStorage() {
        const storedCalendar = localStorage.getItem('scheduledVisitsData');
        const storedBoards = localStorage.getItem('allAvailableBoardsData');
        
        if (storedCalendar) {
            calendarAllocations = JSON.parse(storedCalendar);
            for (const date in calendarAllocations) {
                calendarAllocations[date] = calendarAllocations[date].map(activity => ({
                    ...activity,
                    technicians: activity.technicians || [],
                    shift: activity.shift || shifts[0],
                    module: activity.module || 'N/A'
                }));
            }
        } else {
            calendarAllocations = {};
        }

        if (storedBoards) {
            allAvailableBoards = JSON.parse(storedBoards);
            allAvailableBoards.forEach(board => {
                board.activities = board.activities.map(activity => ({
                    ...activity,
                    module: activity.module || board.id.split('-')[0]
                }));
            });
        } else {
            const { initialBoards } = parseRawSpreadsheetData();
            allAvailableBoards = initialBoards;
        }
    }


    prevMonthBtn.addEventListener('click', () => navigateMonth(-1));
    nextMonthBtn.addEventListener('click', () => navigateMonth(1));

    moduleFilter.addEventListener('change', (event) => {
        filterAndRenderBoards(event.target.value);
    });

    // Modificação AQUI: Botão de impressão agora abre tracking.html para imprimir
    printScheduleBtn.addEventListener('click', () => {
        sessionStorage.setItem('printMode', 'true'); // Sinaliza que é para imprimir
        window.open('tracking.html', '_blank'); // Abre a página de acompanhamento
    });

    integrateSiclaBtn.addEventListener('click', () => {
        updateCalendarDataFromDOM();
        updateBoardDataFromDOM();
        saveScheduleToLocalStorage();

        console.log('INTEGRAÇÃO SICLA: Botão "Integrar SICLA" clicado!');
        console.log('INTEGRAÇÃO SICLA: Dados das agendas (calendarAllocations) para integração/salvamento:', calendarAllocations);
        console.log('INTEGRAÇÃO SICLA: Dados das boards de atividades (allAvailableBoards) para integração/salvamento (incluindo as vazias):', allAvailableBoards);

        alert('Agendas salvas no navegador e prontas para integração com SICLA! (Verifique o console do navegador para os dados que seriam enviados)');
    });


    loadScheduleFromLocalStorage();
    const { uniqueModules: parsedUniqueModules } = parseRawSpreadsheetData();
    populateModuleFilter(parsedUniqueModules);
    populateTechnicianSelector();
    filterAndRenderBoards(moduleFilter.value);
    renderCalendar();
});