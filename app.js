document.addEventListener('DOMContentLoaded', () => {
    const crDataList = document.getElementById('cr-data-list');

    const fetchCRData = async () => {
        try {
            const response = await fetch('http://localhost:5003/Data/cr');
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const data = await response.json();
            
            if (data.length > 0) {
                crDataList.innerHTML = data.map(item => `
                    <div class="p-4 border-b border-gray-200 last:border-b-0">
                        <p><strong>Inventory:</strong> ${item.inventory}</p>
                        <p><strong>Year:</strong> ${item.year}</p>
                        <p><strong>CR:</strong> ${(item.cr * 100).toFixed(2)}%</p>
                        <p><strong>Opened:</strong> ${item.opened}</p>
                        <p><strong>Closed:</strong> ${item.closed}</p>
                        <p><strong>Case Type:</strong> ${item.caseType}</p>
                    </div>
                `).join('');
            } else {
                crDataList.innerHTML = '<p>No CR data found.</p>';
            }

        } catch (error) {
            console.error('Error fetching CR data:', error);
            crDataList.innerHTML = `<p class="text-red-500">Failed to load CR data: ${error.message}</p>`;
        }
    };

    fetchCRData();
});
