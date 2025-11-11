# Extracting Lens Data Using MATLAB

1. Save `MATLABZOSConnection_data.m` and `BatchExportLensData.m` at the same location.
2. Add their folder to MATLAB's PATH.
3. Copy the directory of your lens data into the `lensFolder` variable in `BatchExportLensData.m` line 10.
4. Open Zemax OpticStudio
5. Along the top ribbon of tabs, open `Programming`.
6. Select `Interactive Extension`.
7. In the MATLAB Command Window run: `BatchExportLensData`.


The files should be saved to a new folder called `Exports` within your lens data folder. Becareful about `.ZAR` files, as their lens data may not automatically be added into their respective `.CSV` files. Thus far I have not found a solution for this problem other than manually opening each Zemax Archive file and saving it as a new `.ZMX` file. This process is automatic once you open the archive, but it is tedious nonetheless.