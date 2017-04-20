var exec = require("child_process").exec;

// Package extension
var command = `tfx extension create --overrides-file configs/devssl.json --manifest-globs vss-extension.json --rev-version --no-prompt --json`;
exec(command, (error, stdout) => {
    if (error) {
        console.error(`Could not create package: '${error}'`);
        return;
    }

    console.log(`Package created`);
});