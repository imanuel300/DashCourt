document.addEventListener('DOMContentLoaded', () => {
    const API_BASE_URL = 'DashCourtApi/XLSX'; // Changed to point to the directory containing JSON files

    // Global data storage
    let allCrData = [], allAvgoData = [], allSitData = [], allInv3Data = [];

    // Chart instances
    let movementsChart, crChart, avgDurationChart, hearingsChart;

    // Filter elements
    const yearSelect = document.getElementById('year-filter');
    const courtSelect = document.getElementById('court-filter');
    const caseTypeSelect = document.getElementById('case-type-filter');

    const fetchData = async () => {
        try {
            [allCrData, allAvgoData, allSitData, allInv3Data] = await Promise.all([
                fetch(`${API_BASE_URL}/CR.json`).then(res => res.json()),
                fetch(`${API_BASE_URL}/AVGO.json`).then(res => res.json()),
                fetch(`${API_BASE_URL}/SIT.json`).then(res => res.json()),
                fetch(`${API_BASE_URL}/Inv3.json`).then(res => res.json())
            ]);

            //console.log('Fetched CR Data:', allCrData);
            //console.log('Fetched AVGO Data:', allAvgoData);
            //console.log('Fetched SIT Data:', allSitData);
            //console.log('Fetched Inv3 Data:', allInv3Data);

            populateFilters();
            applyFilters();

        } catch (error) {
            console.error('Error fetching data:', error);
            // Update relevant sections to show error
        }
    };

    const populateFilters = () => {
        // Populate Years
        const allYears = new Set();
        allCrData.forEach(item => allYears.add(item.Year));
        allAvgoData.forEach(item => allYears.add(item.Year));
        allSitData.forEach(item => allYears.add(item.Year));
        allInv3Data.forEach(item => allYears.add(item.Year));

        yearSelect.innerHTML = '<option value="all">שנה</option>' + 
            Array.from(allYears).sort((a, b) => b - a).map(year => `<option value="${year}">${year}</option>`).join('');

        // Populate Court Types (from CR data as an example)
        const allCourts = new Set();
        allCrData.forEach(item => allCourts.add(item.Court));
        courtSelect.innerHTML = '<option value="all">סוג בית משפט</option>' + 
            Array.from(allCourts).sort().map(court => `<option value="${court}">${court}</option>`).join('');

        // Populate Case Types (from CR data as an example)
        const allCaseTypes = new Set();
        allCrData.forEach(item => allCaseTypes.add(item.CaseType));
        caseTypeSelect.innerHTML = '<option value="all">סוג תיק</option>' + 
            Array.from(allCaseTypes).sort().map(caseType => `<option value="${caseType}">${caseType}</option>`).join('');
    };

    const applyFilters = () => {
        const selectedYear = yearSelect.value;
        const selectedCourt = courtSelect.value;
        const selectedCaseType = caseTypeSelect.value;

        const filterData = (data) => {
            return data.filter(item => {
                const matchYear = selectedYear === 'all' || item.Year == selectedYear;
                const matchCourt = selectedCourt === 'all' || item.Court === selectedCourt;
                const matchCaseType = selectedCaseType === 'all' || item.CaseType === selectedCaseType;
                return matchYear && matchCourt && matchCaseType;
            });
        };

        const filteredCrData = filterData(allCrData);
        const filteredAvgoData = filterData(allAvgoData);
        const filteredSitData = filterData(allSitData);
        const filteredInv3Data = filterData(allInv3Data);

        console.log('Filtered CR Data:', filteredCrData);
        console.log('Filtered AVGO Data:', filteredAvgoData);
        console.log('Filtered SIT Data:', filteredSitData);
        console.log('Filtered Inv3 Data:', filteredInv3Data);

        updateKPIs(filteredCrData, filteredAvgoData, filteredSitData, filteredInv3Data);
        renderCharts(filteredCrData, filteredAvgoData, filteredSitData, filteredInv3Data);
    };

    // Add event listeners for filters
    yearSelect.addEventListener('change', applyFilters);
    courtSelect.addEventListener('change', applyFilters);
    caseTypeSelect.addEventListener('change', applyFilters);

    const updateKPIs = (crData, avgoData, sitData, inv3Data) => {
        // Placeholder values for now, as exact calculation logic isn't defined
        // You would typically aggregate data here to get these numbers
        document.getElementById('kpi-judges').innerText = '850';
        document.getElementById('kpi-registrars').innerText = '70';
        document.getElementById('kpi-employees').innerText = '4,500';
    };

    const renderCharts = (crData, avgoData, sitData, inv3Data) => {
        // Destroy existing charts if they exist
        if (movementsChart) movementsChart.destroy();
        if (crChart) crChart.destroy();
        if (avgDurationChart) avgDurationChart.destroy();
        if (hearingsChart) hearingsChart.destroy();

        // Movements Chart (תנועות תיקים)
        const years = [...new Set(crData.map(item => item.Year))].sort();
        const openedData = years.map(year => crData.filter(item => item.Year === year).reduce((sum, item) => sum + item.Opened, 0));
        const closedData = years.map(year => crData.filter(item => item.Year === year).reduce((sum, item) => sum + item.Closed, 0));
        const inventoryData = years.map(year => {
            const yearCrData = crData.filter(item => item.Year === year);
            const totalOpened = yearCrData.reduce((sum, item) => sum + item.Opened, 0);
            const totalClosed = yearCrData.reduce((sum, item) => sum + item.Closed, 0);
            // Assuming inventory is opened - closed for simplicity in mock data context
            return totalOpened - totalClosed; // This might need adjustment based on actual inventory logic
        });

        const movementsCtx = document.getElementById('movementsChart').getContext('2d');
        movementsChart = new Chart(movementsCtx, {
            type: 'bar',
            data: {
                labels: years,
                datasets: [
                    {
                        label: 'נפתחו',
                        data: openedData,
                        backgroundColor: '#36A2EB',
                        datalabels: {
                            anchor: 'end',
                            align: 'top',
                            formatter: (value) => value.toLocaleString() + 'K' // Add K for thousands
                        }
                    },
                    {
                        label: 'נסגרו',
                        data: closedData,
                        backgroundColor: '#FF6384',
                        datalabels: {
                            anchor: 'end',
                            align: 'top',
                            formatter: (value) => value.toLocaleString() + 'K'
                        }
                    },
                    {
                        label: 'מלאי',
                        data: inventoryData,
                        type: 'line',
                        borderColor: '#FFCE56',
                        fill: false,
                        yAxisID: 'y1',
                        datalabels: {
                            anchor: 'end',
                            align: 'top',
                            formatter: (value) => value.toLocaleString()
                        }
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    datalabels: {
                        color: '#000',
                        font: {
                            weight: 'bold'
                        }
                    }
                },
                scales: {
                    x: { 
                        stacked: true 
                    },
                    y: { 
                        stacked: true,
                        beginAtZero: true
                    },
                    y1: {
                        type: 'linear',
                        display: true,
                        position: 'right',
                        grid: {
                            drawOnChartArea: false,
                        }
                    }
                }
            },
            plugins: [ChartDataLabels]
        });

        // CR Chart (קצב סגירת תיקים)
        const crValues = crData.map(item => item.CR);
        const averageCR = crValues.length ? crValues.reduce((sum, cr) => sum + cr, 0) / crValues.length : 0;

        const crCtx = document.getElementById('crChart').getContext('2d');
        crChart = new Chart(crCtx, {
            type: 'doughnut',
            data: {
                labels: ['נסגרו', 'פתוחים'],
                datasets: [{
                    data: [averageCR * 100, (1 - averageCR) * 100],
                    backgroundColor: ['#4CAF50', '#FFC107'],
                    hoverOffset: 4,
                    datalabels: {
                        formatter: (value, ctx) => {
                            return value.toFixed(2) + '%';
                        },
                        color: '#fff',
                        font: {
                            weight: 'bold'
                        }
                    }
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    datalabels: {
                        formatter: (value, ctx) => {
                            return value.toFixed(2) + '%';
                        },
                        color: '#fff',
                        font: {
                            weight: 'bold'
                        }
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                let label = context.label || '';
                                if (label) {
                                    label += ': ';
                                }
                                if (context.parsed !== null) {
                                    label += context.parsed.toFixed(2) + '%';
                                }
                                return label;
                            }
                        }
                    }
                }
            },
            plugins: [ChartDataLabels]
        });

        // Average Case Duration Chart (ממוצע אורך חיי תיק בחודשים)
        const avgDurationDataGrouped = avgoData.reduce((acc, item) => {
            if (!acc[item.CaseType]) {
                acc[item.CaseType] = { totalDays: 0, count: 0 };
            }
            acc[item.CaseType].totalDays += item.AverageDays;
            acc[item.CaseType].count++;
            return acc;
        }, {});

        const caseTypesRaw = Object.keys(avgDurationDataGrouped);
        const avgDaysPerCaseTypeRaw = caseTypesRaw.map(caseType => {
            const data = avgDurationDataGrouped[caseType];
            return data.count ? data.totalDays / data.count : 0;
        });

        // Sort data from highest to lowest
        const sortedData = caseTypesRaw.map((caseType, index) => ({
            caseType: caseType,
            avgDays: avgDaysPerCaseTypeRaw[index]
        })).sort((a, b) => b.avgDays - a.avgDays);

        const caseTypes = sortedData.map(item => item.caseType);
        const avgDaysPerCaseType = sortedData.map(item => item.avgDays.toFixed(1));

        const avgDurationCtx = document.getElementById('avgDurationChart').getContext('2d');
        avgDurationChart = new Chart(avgDurationCtx, {
            type: 'bar',
            data: {
                labels: caseTypes,
                datasets: [
                    {
                        label: 'ממוצע חודשים',
                        data: avgDaysPerCaseType,
                        backgroundColor: '#007CB9', // Changed color to #007CB9
                        // datalabels: { // Removed datalabels to show only on hover
                        //     anchor: 'end',
                        //     align: 'top',
                        //     formatter: (value) => value + ' חודשים'
                        // }
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                let label = context.dataset.label || '';
                                if (label) {
                                    label += ': ';
                                }
                                if (context.parsed.y !== null) {
                                    label += context.parsed.y + ' חודשים';
                                }
                                return label;
                            }
                        }
                    },
                    // datalabels: { // Removed datalabels plugin
                    //     color: '#000',
                    //     font: {
                    //         weight: 'bold'
                    //     }
                    // }
                },
                scales: {
                    y: { beginAtZero: true }
                }
            },
            // plugins: [ChartDataLabels] // Removed ChartDataLabels plugin from here as well
        });

        // Hearings Chart (דיונים שהתקיימו)
        const sitYears = [...new Set(sitData.map(item => item.Year))].sort();
        const hearingsCount = sitYears.map(year => sitData.filter(item => item.Year === year).reduce((sum, item) => sum + item.Hearings, 0));

        const hearingsCtx = document.getElementById('hearingsChart').getContext('2d');
        hearingsChart = new Chart(hearingsCtx, {
            type: 'bar',
            data: {
                labels: sitYears,
                datasets: [
                    {
                        label: 'מספר דיונים',
                        data: hearingsCount,
                        backgroundColor: '#673AB7',
                        datalabels: {
                            anchor: 'end',
                            align: 'top',
                            formatter: (value) => value.toLocaleString()
                        }
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    datalabels: {
                        color: '#000',
                        font: {
                            weight: 'bold'
                        }
                    }
                },
                scales: {
                    y: { beginAtZero: true }
                }
            },
            plugins: [ChartDataLabels]
        });

    };

    fetchData();
});
