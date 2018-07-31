function out = deleteEmptyLines(CNXN,table)
    emptyCells = cellfun('isempty', CNXN);
    isstrprop(CNXN,'wspace');
    cellfun('length',CNXN);
    emptyCells = emptyCells+(cellfun(@sum,isstrprop(CNXN,'alpha'))==0);
    table(all(emptyCells,2),:) = [];
    out = table;
end