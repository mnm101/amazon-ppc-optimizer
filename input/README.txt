Put your two Amazon exports here, renamed to exactly these names:

    keywords.xlsx      <- the keyword/target export   (exports as kt_export_*.xlsx)
    placements.xlsx    <- the placement export        (exports as pl_export_*.xlsx)

Then, from the project folder, run:

    python ppc_optimizer.py --window-days 7

Results are written to outputs/recs.xlsx.

Notes
-----
* --window-days must match the number of days your export actually covers.
  It drives the late-conversion allowance (a 7-day pull has ~15% of its sales
  still un-attributed). Running a 30-day export as 7 would over-cut everything.
* Overwrite these two files after each new export; nothing else needs changing.
* Close outputs/recs.xlsx in Excel before re-running, or the write will fail
  with "Permission denied".
* The .xlsx files in this folder are gitignored - your account data is not
  committed.
* To use different files just once, pass them explicitly instead:
      python ppc_optimizer.py SOME_KW.xlsx SOME_PL.xlsx -o outputs/other.xlsx
