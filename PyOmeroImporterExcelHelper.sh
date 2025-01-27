#!/bin/bash
curDir="$( cd -- "$(dirname "$0")" >/dev/null 2>&1 ; pwd -P )"
script="${curDir}/fetch_images.py"
excelFilePath="$1"
isDebug="$2"
result= false


activateScript="${curDir}/.venv/bin/activate"

if [[ ! -f $activateScript ]]; then
    activateScript="${curDir}/venv3/bin/activate"
fi
echo "Activating venv ${activateScript}"
source $activateScript

if [[ -n "$excelFilePath" ]]; then
    eval "excelFilePath=$excelFilePath"
    if [[ -f "$excelFilePath" ]]; then
        echo "Running python for ${excelFilePath}"
        python3 $script $excelFilePath
    fi
fi

echo "Deactivating venv"
deactivate