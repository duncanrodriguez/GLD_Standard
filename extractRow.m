function outarray = extractRow( cellarray, rowin, outarray)
rowout = length(outarray{1})+1;
for i = 1:length(cellarray)
    outarray{i}(rowout,1) = cellarray{i}(rowin,1);
end
end