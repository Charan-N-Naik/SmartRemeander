// Fetch and display recent reminders
fetch('/recent')
    .then(response => response.json())
    .then(data => {
        const container = document.getElementById('recent-reminders');
        if (data.length === 0) {
            container.innerHTML = '<p>No reminders yet.</p>';
        } else {
            container.innerHTML = data.map(item => 
                `<p><strong>${item.index}:</strong> ${item.dateTime} - ${item.activity}</p>`
            ).join('');
        }
    })
    .catch(error => {
        document.getElementById('recent-reminders').innerHTML = '<p>Error loading recent reminders.</p>';
    });

// Form submission
document.getElementById('reminderForm').addEventListener('submit', function(event) {
    event.preventDefault();
    const date = document.getElementById('date').value;
    const time = document.getElementById('time').value;
    const task = document.getElementById('task').value;
    const email = document.getElementById('email').value;
    
    fetch('/add-reminder', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ date, time, task, email })
    })
    .then(response => {
        if (response.redirected) {
            window.location.href = response.url;
        } else {
            return response.text();
        }
    })
    .then(data => {
        if (data && data.includes('Error')) {
            alert(data);
        }
    })
    .catch(error => console.error('Error:', error));
});
