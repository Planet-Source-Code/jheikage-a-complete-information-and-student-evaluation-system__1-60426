SELECT COUNT(Moccup) AS Total, Moccup, Course, Inx.Yr
FROM NAMES INNER JOIN
        (SELECT SubjectsEnrolled.yr, subjectsenrolled.lnam, 
           subjectsenrolled.fnam, subjectsenrolled.mnam, 
           subjectsenrolled.sy, subjectsenrolled.sem
      FROM subjectsenrolled
      GROUP BY lnam, fnam, mnam, sy, sem, yr) AS inx ON 
    inx.lnam = names.lnam AND inx.fnam = names.fnam AND 
    inx.mnam = names.mnam
GROUP BY moccup, Course, Inx.yr

'Sample COmmand for grouping ito pare....iehjsuckers