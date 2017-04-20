var exec = require("child_process").exec;

// Package extension
var command = `tfx extension create --overrides-file ../configs/release.json --manifest-globs vss-extension.json --rev-version --no-prompt --json`;
exec(command, { cwd: 'dist/' }, (error, stdout) => {
    if (error) {
        console.error(`Could not create package: '${error}'`);
        return;
    }

    console.log(`Package created`);
});