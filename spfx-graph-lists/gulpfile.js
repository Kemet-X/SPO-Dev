const { spawn } = require('child_process');

function run(command, args) {
  return (done) => {
    const child = spawn(command, args, {
      stdio: 'inherit',
      shell: process.platform === 'win32'
    });

    child.on('close', (code) => {
      if (code === 0) {
        done();
        return;
      }

      done(new Error(`${command} ${args.join(' ')} exited with code ${code}`));
    });
  };
}

exports.build = run('npm', ['run', 'build']);
exports.serve = run('npm', ['run', 'start']);
exports.clean = run('npm', ['run', 'clean']);
