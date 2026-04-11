#!/bin/bash

set -eu


## works both under bash and sh
SCRIPT_DIR=$(dirname "$(readlink -f "$0")")


## leave empty to resolve by PATH
## set variable in case of virtual environment - it will check and ensure that tool installation is added to pyproject.toml
COMMAND_PATH=""
if [ ! -z ${VIRTUAL_ENV+x} ]; then
    ## Python virtual environment detected -- use command from venv
    COMMAND_PATH="${VIRTUAL_ENV}/bin/"
fi


src_dir="$SCRIPT_DIR/../src"

examples_dir="$SCRIPT_DIR/../examples"
if [ ! -d "$examples_dir" ]; then
    examples_dir=""
fi


## ============================================

## D100: Missing docstring in public module
## D101: Missing docstring in public class
## D102: Missing docstring in public method
## D103: Missing docstring in public function
## D104: Missing docstring in public package
## D105: Missing docstring in magic method
## D107 Missing docstring in __init__
ignore_errors=D100,D101,D102,D103,D104,D105,D107


echo "running pydocstyle"
python3 -m pydocstyle --count --convention=numpy --add-ignore="$ignore_errors" "$src_dir" "$examples_dir" "$SCRIPT_DIR"
# pydocstyle --count --ignore=$ignore_errors $src_dir
exit_code=$?

if [ $exit_code -ne 0 ]; then
    exit $exit_code
fi

echo "pydocstyle -- no warnings found"


## ============================================

disabled() {
#
# there is problem with dargling - there is no possiblitiy
# to skip certain checks through command-line
#

echo "running darglint"

src_files=$(find "$src_dir" -name "*.py")
examples_files=$(find "$examples_dir" -name "*.py")
local_files=$(find "$SCRIPT_DIR" -name "*.py")

"${COMMAND_PATH}"darglint "${src_files[@]}" "${examples_files[@]}" "${local_files[@]}"
exit_code=$?

if [ $exit_code -ne 0 ]; then
    exit $exit_code
fi

echo "darglint -- no warnings found"
}
