function [ error ] = calcerror( known, measured )
error = abs(known-measured)/known*100;
end

