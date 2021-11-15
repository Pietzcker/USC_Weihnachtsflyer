SELECT DISTINCTROW p.Nummer, p.Einrichtung, p.Titel, p.Vorname, p.Name, p.[Straße/Postfach], p.PLZ, p.Ort, p.Land, a.Briefanrede AS Briefanrede, g.Adressanrede AS Adressanrede, p.Geburtsdatum, "Mitglied" AS Mitglied
FROM ((COMPBereiche b INNER JOIN COMPersonen p ON b.[Person (Nr)] = p.Nummer)
INNER JOIN COMBriefanreden a ON a.Nummer = p.[Briefanrede (Nr)])
INNER JOIN COMAdressanreden g on g.Nummer = p.[Adressanrede (Nr)]
WHERE (
	(b.[Bereich (Nr)] >= 1 AND b.[Bereich (Nr)] <= 5) 
	OR
	b.[Bereich (Nr)] = 10 OR b.[Bereich (Nr)] = 20 OR b.[Bereich (Nr)] = 30 OR b.[Bereich (Nr)] = 35
)
AND p.[Straße/Postfach] IS NOT NULL

UNION

SELECT DISTINCTROW p.Nummer, p.Einrichtung, p.Titel, p.Vorname, p.Name, p.[Straße/Postfach], p.PLZ, p.Ort, p.Land, a.Briefanrede AS Briefanrede, g.Adressanrede AS Adressanrede, p.Geburtsdatum, "Schnupperer" AS Mitglied
FROM ((COMPBereiche b INNER JOIN COMPersonen p ON b.[Person (Nr)] = p.Nummer)
INNER JOIN COMBriefanreden a ON a.Nummer = p.[Briefanrede (Nr)])
INNER JOIN COMAdressanreden g on g.Nummer = p.[Adressanrede (Nr)]
WHERE (
	(b.[Bereich (Nr)] >= 6 AND b.[Bereich (Nr)] <= 9) 
	OR
	b.[Bereich (Nr)] = 11
)
AND b.Bis IS NULL AND b.Von <= Date()
AND p.[Straße/Postfach] IS NOT NULL

UNION

SELECT DISTINCTROW p.Nummer, p.Einrichtung, p.Titel, p.Vorname, p.Name, p.[Straße/Postfach], p.PLZ, p.Ort, p.Land, a.Briefanrede AS Briefanrede, g.Adressanrede AS Adressanrede, p.Geburtsdatum, "Flyer" AS Mitglied
FROM ((COMPRubriken b INNER JOIN COMPersonen p ON b.[Person (Nr)] = p.Nummer)
INNER JOIN COMBriefanreden a ON a.Nummer = p.[Briefanrede (Nr)])
INNER JOIN COMAdressanreden g on g.Nummer = p.[Adressanrede (Nr)]
WHERE b.[Rubrik (Nr)] = 24
AND p.[Straße/Postfach] IS NOT NULL