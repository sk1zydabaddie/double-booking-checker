import * as React from "react";
import { Button, makeStyles, Text, Card, CardHeader } from "@fluentui/react-components";
import { Search24Regular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  root: {
    padding: "20px",
  },
  card: {
    marginTop: "20px",
  },
  button: {
    marginTop: "10px",
  },
  results: {
    marginTop: "20px",
  },
});

const BookingChecker = () => {
  const styles = useStyles();
  const [isScanning, setIsScanning] = React.useState(false);
  const [results, setResults] = React.useState(null);
  const [isOfficeInitialized, setIsOfficeInitialized] = React.useState(false);

  React.useEffect(() => {
    Office.onReady(() => {
      setIsOfficeInitialized(true);
    });
  }, []);

  const handleScan = async () => {
    if (!isOfficeInitialized) {
      setResults("Error: Office.js is not initialized yet. Please try again in a moment.");
      return;
    }

    setIsScanning(true);
    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load(["values", "columnCount", "rowCount"]);
        
        await context.sync();

        const values = usedRange.values;
        const headers = values[0];

        // Find date/time related columns
        const dateColumns = [];
        headers.forEach((header, index) => {
          const headerStr = String(header).toLowerCase();
          if (headerStr.includes('date') || headerStr.includes('time') || 
              headerStr.includes('start') || headerStr.includes('end')) {
            dateColumns.push(index);
          }
        });

        if (dateColumns.length < 2) {
          setResults("Could not find enough date/time columns. Please ensure your sheet has columns for start and end times.");
          return;
        }

        // Assume the first two date columns are start and end times
        const startCol = dateColumns[0];
        const endCol = dateColumns[1];

        // Process each row and look for conflicts
        const conflicts = [];
        for (let i = 1; i < values.length; i++) {
          const row1Start = new Date(values[i][startCol]);
          const row1End = new Date(values[i][endCol]);

          if (isNaN(row1Start.getTime()) || isNaN(row1End.getTime())) continue;

          for (let j = i + 1; j < values.length; j++) {
            const row2Start = new Date(values[j][startCol]);
            const row2End = new Date(values[j][endCol]);

            if (isNaN(row2Start.getTime()) || isNaN(row2End.getTime())) continue;

            // Check for overlap
            if (row1Start < row2End && row2Start < row1End) {
              conflicts.push({
                booking1: `Row ${i + 1}: ${row1Start.toLocaleString()} - ${row1End.toLocaleString()}`,
                booking2: `Row ${j + 1}: ${row2Start.toLocaleString()} - ${row2End.toLocaleString()}`
              });
            }
          }
        }

        if (conflicts.length === 0) {
          setResults("No booking conflicts found.");
        } else {
          const conflictMessages = conflicts.map(conflict => 
            `Conflict between:\n${conflict.booking1}\n${conflict.booking2}`
          );
          setResults(`Found ${conflicts.length} booking conflict(s):\n\n${conflictMessages.join('\n\n')}`);
        }
      });
    } catch (error) {
      setResults(`Error: ${error.message}`);
    } finally {
      setIsScanning(false);
    }
  };

  return (
    <div className={styles.root}>
      <Card className={styles.card}>
        <CardHeader header={<Text weight="semibold">Double Booking Checker</Text>} />
        <Text>Click the button below to scan your schedule for any double bookings.</Text>
        <Button
          appearance="primary"
          icon={<Search24Regular />}
          onClick={handleScan}
          disabled={isScanning || !isOfficeInitialized}
          className={styles.button}
        >
          {isScanning ? "Scanning..." : "Scan for Conflicts"}
        </Button>
      </Card>
      {results && (
        <Card className={styles.results}>
          <CardHeader header={<Text weight="semibold">Results</Text>} />
          <Text>{results}</Text>
        </Card>
      )}
    </div>
  );
};

export default BookingChecker;