function [output,PF] = tokW(Voltage,PF,leadlag,power,ratedUnits)
    for i = 1:length(PF)
        switch ratedUnits(i)
            case 0
                output(i,1) = PF(i)*power(i);
            case 1
                output(i,1) = PF(i)*power(i)*1000;
            case 2
                output(i,1) = PF(i)*power(i)*Voltage(i)/1000;
            case 3
                output(i,1) = power(i)*0.7457;
            case 4
                output(i,1) = power(i);
            case 5
                output(i,1) = power(i)*1000;
            case 7
                output(i,1) = power(i);
            case 8
                output(i,1) = power(i)*1000;
        end
        switch leadlag(i)
            case 1
                PF(i) = -PF(i);
        end
    end
end

