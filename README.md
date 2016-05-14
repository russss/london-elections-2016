# London 2016 Election Data

This repository contains processed ward-level voting stats for the 2016 London
mayoral and assembly elections.

* `results.xlsx` is the published ward-level data, in an awful
spreadsheet.
* `parse.py` parses this spreadsheet to somewhat nicer CSV files.
* The generated CSV files are in `./data`

Wards have been matched with the [GSS area
codes](https://en.wikipedia.org/wiki/ONS_coding_system#Current_GSS_coding_system)
where possible. One exception to this is the City of London, which
doesn't seem have votes counted at ward level in London elections.

Postal votes are counted as separate pseudo-wards.
