// Initialize the Microsoft Teams SDK
microsoftTeams.initialize();

// Function to show a dialog
function showDialog() {
    microsoftTeams.authentication.authenticate({
        url: window.location.origin + '/auth-start.html',
        successCallback: function (result) {
            // Authentication succeeded
            console.log('Authentication succeeded:', result);
            microsoftTeams.authentication.notifySuccess(result);
        },
        failureCallback: function (reason) {
            // Authentication failed
            console.error('Authentication failed:', reason);
            microsoftTeams.authentication.notifyFailure(reason);
        }
    });
}

// Add a click event listener to the "Show Dialog" button
document.getElementById('showDialogButton').addEventListener('click', showDialog);
