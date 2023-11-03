microsoftTeams.initialize();

function displayTasks() {
    // Assuming you have a backend API endpoint to fetch tasks
    fetch('https://api.example.com/tasks')
        .then(response => response.json())
        .then(tasks => {
            const taskList = document.getElementById('taskList');
            // Clear existing task list
            taskList.innerHTML = '';

            // Populate task list with fetched tasks
            tasks.forEach(task => {
                const taskItem = document.createElement('div');
                taskItem.textContent = task.name; // Assuming each task object has a 'name' property
                taskList.appendChild(taskItem);
            });
        })
        .catch(error => {
            console.error('Error fetching tasks:', error);
        });
}
