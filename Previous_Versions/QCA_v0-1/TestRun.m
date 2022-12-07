function [capsules,bodies] = TestRun(filename, cap_min, cap_max, cap_sens, bod_min, bod_max, bod_sens)

rgb = imread(filename);
%In case they are not yet greyscale compatible
if size(rgb,3) == 3
    rgb = rgb2gray(rgb);
end

figure;
%Generate a visible and well-contrasted image
adjust = imadjust(rgb);

%Automatic thresholding
ld = double(rgb);
Av = mean2(ld);
Sig = std2(ld);
%Use the average value + standard deviation as threshold
auto_bin = rgb > (Av + Sig);

%identify the circle centers and radii
[centers, radii] = imfindcircles(auto_bin,[cap_min,cap_max],'Sensitivity',cap_sens,'ObjectPolarity','bright');
[b_centers, b_radii] = imfindcircles(rgb,[bod_min, bod_max],'ObjectPolarity','dark','Sensitivity',bod_sens);
imshow(adjust);
viscircles(centers,radii,'EdgeColor','g');
viscircles(b_centers,b_radii,'EdgeColor','b');
saveas(gcf,'TestRun.jpeg');
close

capsules = cell2mat(num2cell([centers, radii]));
bodies = cell2mat(num2cell([b_centers, b_radii]));
end