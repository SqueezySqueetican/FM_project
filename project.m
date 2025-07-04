clear all;
clc;

[num, txt, raw] = xlsread('PumpsData.xlsx');

q = input('Input flow rate (q [m^3/h]): ');
h = input('Input head (h [m]): ');

Error_min = inf;

for i = 1:6:60
    diameters = num(~isnan(num(:, i)), i);
    formulaText = txt(1 + find(diameters), i + 1);
    pumpName = txt{1, i};
    
    for d = 1:length(diameters)
        hqFormula = formulaText{d};
        q_max_hq = num(d, i + 2);

        if ischar(hqFormula) && q <= q_max_hq
            try
                y = eval(strrep(hqFormula, 'x', num2str(q)));

                if y <= h
                    error = h - y;

                    if error < Error_min
                        Error_min = error;
                        bestPumpName = pumpName;
                        bestImpeller = diameters(d);

                        powerFormula = raw{d+1, i+3};
                        q_min = raw{d+1, i+4};
                        q_max = raw{d+1, i+5};
  
                    end
                end
            end
        end
    end
end

if isfinite(Error_min)
    fprintf('\nRecomended pump: "%s"\n', bestPumpName);
    fprintf('Recommended impeller diameter: %.1f m\n', bestImpeller);
    fprintf('Head error: %.3f meters\n', Error_min);

    if ischar(powerFormula) && isnumeric(q_min) && isnumeric(q_max)
        if q >= q_min && q <= q_max
            p = eval(strrep(powerFormula, 'x', num2str(q)));
            fprintf('Estimated power consumption: %.2f W \n', p);
        else
            fprintf('Flow rate %.2f is out of allowed range for power formula [%.2f – %.2f].\n', q, q_min_power, q_max_power);
        end
    else
        fprintf('Power formula or flow limits missing or invalid.\n');
    end
else
    fprintf('\nNo suitable pump found within head constraint.\n');
end