export const PROCESSOR_TYPES = {
    FILTER: 'filter',
    TRANSFORM: 'transform',
    DELETE_COLUMN: 'delete_column',
    CALCULATE: 'calculate',
};

export const OPERATORS = {
    EQUALS: '==',
    CONTAINS: 'contains',
    GREATER_THAN: '>',
    LESS_THAN: '<',
};

export function applyAutomation(data, headers, rules) {
    let processedData = [...data];
    let processedHeaders = [...headers];

    for (const rule of rules) {
        switch (rule.type) {
            case PROCESSOR_TYPES.FILTER:
                // Dynamic lookup of index because columns might observe shifts
                const filterColIndex = processedHeaders.indexOf(rule.columnName);
                if (filterColIndex === -1) break; // Column not found (maybe deleted)

                processedData = processedData.filter(row => {
                    const cellValue = row[filterColIndex];
                    const targetValue = rule.value;

                    // Basic type conversion for numbers
                    const numCell = parseFloat(cellValue);
                    const numTarget = parseFloat(targetValue);
                    const isNum = !isNaN(numCell) && !isNaN(numTarget);

                    switch (rule.operator) {
                        case OPERATORS.EQUALS:
                            return String(cellValue).toLowerCase() === String(targetValue).toLowerCase();
                        case OPERATORS.CONTAINS:
                            return String(cellValue).toLowerCase().includes(String(targetValue).toLowerCase());
                        case OPERATORS.GREATER_THAN:
                            return isNum ? numCell > numTarget : cellValue > targetValue;
                        case OPERATORS.LESS_THAN:
                            return isNum ? numCell < numTarget : cellValue < targetValue;
                        default:
                            return true;
                    }
                });
                break;

            case PROCESSOR_TYPES.DELETE_COLUMN:
                // We actually need to remove it from headers first to keep indices synced for subsequent steps?
                // Actually, deleting columns makes index-based tracking hard for subsequent rules if we rely on initial indices.
                // For MVP, we'll just reconstruct the row.
                // Better strategy: We just mark the index as null or filter it out at the VERY END.
                // But for chaining, we need immediate feedback.
                // Let's simpler approach: Use column NAMES for tracking, not indices.

                // Find current index of the column to delete
                const colIndexToDelete = processedHeaders.indexOf(rule.columnName);
                if (colIndexToDelete > -1) {
                    processedHeaders.splice(colIndexToDelete, 1);
                    processedData = processedData.map(row => {
                        const newRow = [...row];
                        newRow.splice(colIndexToDelete, 1);
                        return newRow;
                    });
                }
                break;

            case PROCESSOR_TYPES.TRANSFORM:
                const colIndexToTrans = processedHeaders.indexOf(rule.columnName);
                if (colIndexToTrans > -1) {
                    processedData = processedData.map(row => {
                        const newRow = [...row];
                        if (rule.transformation === 'uppercase') {
                            newRow[colIndexToTrans] = String(newRow[colIndexToTrans] || '').toUpperCase();
                        } else if (rule.transformation === 'lowercase') {
                            newRow[colIndexToTrans] = String(newRow[colIndexToTrans] || '').toLowerCase();
                        } else if (rule.transformation === 'trim') {
                            newRow[colIndexToTrans] = String(newRow[colIndexToTrans] || '').trim();
                        }
                        return newRow;
                    });
                }
                break;
        }
    }

    return { data: processedData, headers: processedHeaders };
}
