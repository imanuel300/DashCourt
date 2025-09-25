document.addEventListener('DOMContentLoaded', () => {
    const API_BASE_URL = 'http://localhost:5003/Data';

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
                fetch(`${API_BASE_URL}/cr`).then(res => res.json()),
                fetch(`${API_BASE_URL}/avgo`).then(res => res.json()),
                fetch(`${API_BASE_URL}/sit`).then(res => res.json()),
                fetch(`${API_BASE_URL}/inv3`).then(res => res.json())
            ]);

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
        allCrData.forEach(item => allYears.add(item.year));
        allAvgoData.forEach(item => allYears.add(item.year));
        allSitData.forEach(item => allYears.add(item.year));
        allInv3Data.forEach(item => allYears.add(item.year));

        yearSelect.innerHTML = '<option value="all">שנה</option>' + 
            Array.from(allYears).sort((a, b) => b - a).map(year => `<option value="${year}">${year}</option>`).join('');

        // Populate Court Types (from CR data as an example)
        const allCourts = new Set();
        allCrData.forEach(item => allCourts.add(item.court));
        courtSelect.innerHTML = '<option value="all">סוג בית משפט</option>' + 
            Array.from(allCourts).sort().map(court => `<option value="${court}">${court}</option>`).join('');

        // Populate Case Types (from CR data as an example)
        const allCaseTypes = new Set();
        allCrData.forEach(item => allCaseTypes.add(item.caseType));
        caseTypeSelect.innerHTML = '<option value="all">סוג תיק</option>' + 
            Array.from(allCaseTypes).sort().map(caseType => `<option value="${caseType}">${caseType}</option>`).join('');
    };

    const applyFilters = () => {
        const selectedYear = yearSelect.value;
        const selectedCourt = courtSelect.value;
        const selectedCaseType = caseTypeSelect.value;

        const filterData = (data) => {
            return data.filter(item => {
                const matchYear = selectedYear === 'all' || item.year == selectedYear;
                const matchCourt = selectedCourt === 'all' || item.court === selectedCourt;
                const matchCaseType = selectedCaseType === 'all' || item.caseType === selectedCaseType;
                return matchYear && matchCourt && matchCaseType;
            });
        };

        const filteredCrData = filterData(allCrData);
        const filteredAvgoData = filterData(allAvgoData);
        const filteredSitData = filterData(allSitData);
        const filteredInv3Data = filterData(allInv3Data);

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
        const years = [...new Set(crData.map(item => item.year))].sort();
        const openedData = years.map(year => crData.filter(item => item.year === year).reduce((sum, item) => sum + item.opened, 0));
        const closedData = years.map(year => crData.filter(item => item.year === year).reduce((sum, item) => sum + item.closed, 0));

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
                    },
                    {
                        label: 'נסגרו',
                        data: closedData,
                        backgroundColor: '#FF6384',
                    }
                ]
            },
            options: {
                responsive: true,
                scales: {
                    x: { stacked: true },
                    y: { stacked: true }
                }
            }
        });

        // CR Chart (קצב סגירת תיקים)
        const crValues = crData.map(item => item.cr);
        const averageCR = crValues.length ? crValues.reduce((sum, cr) => sum + cr, 0) / crValues.length : 0;

        const crCtx = document.getElementById('crChart').getContext('2d');
        crChart = new Chart(crCtx, {
            type: 'doughnut',
            data: {
                labels: ['נסגרו', 'פתוחים'],
                datasets: [{
                    data: [averageCR * 100, (1 - averageCR) * 100],
                    backgroundColor: ['#4CAF50', '#FFC107'],
                    hoverOffset: 4
                }]
            },
            options: {
                responsive: true,
                plugins: {
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
            }
        });

        // Average Case Duration Chart (ממוצע אורך חיי תיק בחודשים)
        const avgYears = [...new Set(avgoData.map(item => item.year))].sort();
        const avgDays = avgYears.map(year => {
            const yearData = avgoData.filter(item => item.year === year);
            const sum = yearData.reduce((s, item) => s + item.averageDays, 0);
            return yearData.length ? sum / yearData.length : 0;
        });

        const avgDurationCtx = document.getElementById('avgDurationChart').getContext('2d');
        avgDurationChart = new Chart(avgDurationCtx, {
            type: 'bar',
            data: {
                labels: avgYears,
                datasets: [
                    {
                        label: 'ממוצע ימים',
                        data: avgDays,
                        backgroundColor: '#FF9800',
                    }
                ]
            },
            options: {
                responsive: true,
                scales: {
                    y: { beginAtZero: true }
                }
            }
        });

        // Hearings Chart (דיונים שהתקיימו)
        const sitYears = [...new Set(sitData.map(item => item.year))].sort();
        const hearingsCount = sitYears.map(year => sitData.filter(item => item.year === year).reduce((sum, item) => sum + item.hearings, 0));

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
                    }
                ]
            },
            options: {
                responsive: true,
                scales: {
                    y: { beginAtZero: true }
                }
            }
        });

    };

    fetchData();
});
