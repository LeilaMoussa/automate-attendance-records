# automate-attendance-records
Easily update attendance records on Google Sheets from automatically generated CSV files.
Cannot be used with Excel files whose extensions were changed -- must be an actual CSV as generated by MS Teams "Download Attendance List" feature.
Any help/feedback is super appreciated.

Todos:
- [x] Decrease the number or at least the frequency of API calls to stay within the quota: consider getting all values once and writing once.
- [x] Active members
- [x] Popular events (needs testing)
- [ ] Tutoring
- [x] Fix encoding issues (somewhat)
- [ ] Hide constants in a separate module.

Future work:
- [ ] Typos in names: what differentiates a typo from a slighly different name? ==> Fuzzy search!
