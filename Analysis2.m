function [] = Analysis2(filename, cap_min, cap_max, cap_sens, bod_min, bod_max, bod_sens, image_dir)
    %The extra variables below are just to keep the middle one
    [pathstr,basename,ext] = fileparts(filename);
    %for every file we will do the counting and the saving
    rgb = imread(filename);
    %In case they are not yet greyscale compatible
    if size(rgb,3) == 3
        rgb = rgb2gray(rgb);
    end

    %make the picture and save it to the new folder
    figure;
    %Generate a pretty image for the output file
    adjust = imadjust(rgb);

    %AUTOMATIC THRESHOLDING
    ld = double(rgb);
    Av = mean2(ld);
    Sig = std2(ld);
    %Use the average value + standard deviation as threshold
    auto_bin = rgb > (Av + Sig);

    %identify the circle centers and radii
    [centers, radii] = imfindcircles(auto_bin,[cap_min,cap_max],'Sensitivity',cap_sens,'ObjectPolarity','bright');
    [b_centers, b_radii] = imfindcircles(rgb,[bod_min,bod_max],'ObjectPolarity','dark','Sensitivity',bod_sens);
    imshow(adjust);
    viscircles(centers,radii,'EdgeColor','g');
    viscircles(b_centers,b_radii,'EdgeColor','b');

    mkdir(image_dir,'Processed Results')
    basename2 = [basename, '-processed.jpeg'];
    fullFileName = fullfile(image_dir,'/Processed Results', basename2);
    saveas(gcf, fullFileName);
    close

    %format and output the results to excel
    %sheets named after samples
    A = {'Capsule Center - X','Capsule Center - Y','Capsule Radius (pixel)'};
    B = num2cell([centers, radii]);
    C = {'Body Center - X','Body Center - Y','Body Radius (pixel)'};
    D = num2cell([b_centers, b_radii]);
    %now I want to put the matricies side by side
    m1 = [A;B];
    m2 = [C;D];
    [n,m] = size(m1);
    [n2,m3] = size(m2);
    nn = max(n,n2);
    %this is in case there are a different amount detected
    output = cell(nn,m+m3);
    output(1:n,1:m) = m1;
    output(1:n2,m+1:end) = m2;

    fullFileName2 = fullfile(image_dir, '/RawOutput.xlsx');
    xlswrite(fullFileName2,output,basename);
end