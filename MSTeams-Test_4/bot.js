const { TeamsActivityHandler } = require('botbuilder');

class MyBot extends TeamsActivityHandler {
    constructor() {
        super();

        this.onMessage(async (context, next) => {
            // When a user sends a message, reply with a simple message
            await context.sendActivity('Hello from your Teams bot!');
            await next();
        });
    }
}

module.exports.MyBot = MyBot;
