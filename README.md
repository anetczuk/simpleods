# simpleods

Simple wrapper objects for `odfpy` package allowing easy navigation through spreadsheet with some handy methods.


## Features

- adding, moving, copying and removing rows and cells, 
- getting rows and cells by index,
- printing cells spreadsheet coordinates,
- sorting rows,
- removing empty rows,
- expanding repeated rows and cells


## Installation

Installation of package for use can be done by:
  - to install package from downloaded ZIP file execute: `pip3 install --user -I file:simpleods-master.zip`
  - to install package directly from GitHub execute: `pip3 install --user -I git+https://github.com/anetczuk/simpleods.git`
  - installation from local repository root directory: `pip3 install --user .`

To uninstall run: `pip3 uninstall simpleods`

To install project under virtual environment use `tools/installvenv.sh`.

Development installation is covered in [Development](#development) section.


## Development

Project contains several tools and features that facilitate development and maintenance of the project.

In case of pull requests please run `process-all.sh --release` before the request to check installation, run tests and
perform source code static analysis.


### Installation

Installation for development with configuration of virtual environment:
  - `tools/installvenv.sh --dev` to install dependencies, the package in editable mode and install development tooling.

Installation for development without venv:
  - `src/install-app.sh --dev` to install dependencies, the package in editable mode and install development tooling.

Virtual environmnt can be also configured manually by:
  - `python3 -m venv .venv`
  - `source .venv/bin/activate`
  - `python -m pip install --upgrade pip`
  - `src/install-app.sh --dev` to install dependencies, the package in editable mode and install development tooling
or `python -m pip install -e '.[dev]'` to install project by hand.

There is also possibility to work on the project without installation. In this case project will run from local repository 
directory. This configuration requires installation of dependencies: `./src/install-deps.sh --dev`.


### Running tests

To run tests execute `src/testtestsimpleods/runtests.py`. It can be run with code profiling 
and code coverage options.


### Tools scripts

In *tools* directory there can be found following helper scripts:
- `codecheck.sh` -- static code check using several tools with defined set of rules
- `doccheck.sh` -- run `pydocstyle` with defined configuration
- `mdcheck.sh` -- check links in Markdown files
- `typecheck.sh` -- run `mypy` with defined configuration
- `checkall.sh` -- execute *check* scripts all at once
- `profiler.sh` -- profile Python scripts
- `coverage.sh` -- measure code coverate
- `notrailingwhitespaces.sh* -- as name states removes trailing whitespaces from _*.py*_ files
- `rmpyc.sh` -- remove all _*.pyc_ files

Those scripts can be run also from within virtual environment.


## References

- https://github.com/eea/odfpy


## License

```
BSD 3-Clause License

Copyright (c) 2026, Arkadiusz Netczuk <dev.arnet@gmail.com>

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

3. Neither the name of the copyright holder nor the names of its
   contributors may be used to endorse or promote products derived from
   this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
```
