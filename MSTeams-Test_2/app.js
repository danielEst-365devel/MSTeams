// Initialize the Microsoft Teams SDK
microsoftTeams.initialize();

// Function to update the greeting
function updateGreeting() {
    microsoftTeams.getContext(function (context) {
        const userName = context.userDisplayName || 'User';
        document.getElementById('greeting').textContent = `Hello, ${userName}!`;
    });
}

// Add a click event listener to the "Update Greeting" button
document.getElementById('updateGreeting').addEventListener('click', updateGreeting);

// Call updateGreeting when the page loads
updateGreeting();
