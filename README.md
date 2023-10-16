SELECT
    ID,
    Description AS OriginalString,
    SUBSTRING(
        Description,
        CHARINDEX(',', Description, CHARINDEX(',', Description, CHARINDEX(',', Description) + 1) + 1) + 1,
        CHARINDEX(',', Description, CHARINDEX(',', Description, CHARINDEX(',', Description) + 1) + 1) -
        CHARINDEX(',', Description, CHARINDEX(',', Description, CHARINDEX(',', Description) + 1) + 1) - 1
    ) AS ExtractedValue
FROM SampleData;
