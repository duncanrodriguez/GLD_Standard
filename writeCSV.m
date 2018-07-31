function [] = writeCSV(name,mat,dir,titleCol)
currfile = [name,'.csv'];
Table = toTable(mat,titleCol);
writeFile = fopen(currfile,'w');
writetable(Table,currfile);
fclose(writeFile);
try
    copyfile(currfile,dir);
    delete(currfile);
catch
    disp(['Notice: The "',currfile, '" file is in ', pwd,' since it was not succesfully copied to ',  dir,' .']);
end
end