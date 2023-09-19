const spawn = require("child_process").spawn;
const pythonProcess = spawn('python',["./manage.py", 'runcrons'], {
    stdio: 'inherit'
})

pythonProcess.on('exit', ()=>{
    process.exit(1)
})
pythonProcess.on('close', ()=>{
    process.exit(1)
})