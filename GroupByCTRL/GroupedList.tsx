import * as React from "react";
import { useState, useEffect } from "react";
import { GroupedList, IGroup } from "@fluentui/react/lib/GroupedList";
import {
  IColumn,
  DetailsRow,
  DetailsHeader,
} from "@fluentui/react/lib/DetailsList";
import {
  Selection,
  SelectionMode,
  SelectionZone,
} from "@fluentui/react/lib/Selection";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { useConst } from "@fluentui/react-hooks";

export type Dataset = ComponentFramework.PropertyTypes.DataSet;

export interface IGroupedListProps {
  dataset: Dataset;
  entityName: string;
  context: ComponentFramework.Context<IInputs>;
  currentPage: number;
  loadPrevPage: () => void;
  loadNextPage: () => void;
}

export type ILookupValue = ComponentFramework.LookupValue;
export type IEntityReference = ComponentFramework.EntityReference;

export interface DynamicItem {
  [key: string]:
    | string
    | Date
    | number
    | number[]
    | boolean
    | IEntityReference
    | IEntityReference[]
    | ILookupValue
    | ILookupValue[]
    | undefined;
}

//#region Making groups
export function makeGroups(
  dataset: Dataset,
  level1: string,
  level2: string
): { groups: IGroup[]; items: DynamicItem[] } {
  const groups: IGroup[] = [];
  const items: DynamicItem[] = [];
  let startIndex = 0;

  if (!dataset || !Array.isArray(dataset.sortedRecordIds) || !dataset.records) {
    return { groups, items };
  }

  // Create unique groups for level1
  const uniqueLevel1Values = new Set<string>();
  dataset.sortedRecordIds.forEach((id) => {
    const record = dataset.records[id];
    const value1 = record.getFormattedValue(level1);
    if (value1) {
      uniqueLevel1Values.add(value1);
    }
  });

  // Iterate over unique level1 values and create subgroups for level2
  uniqueLevel1Values.forEach((value1) => {
    const level2Groups: IGroup[] = [];
    const uniqueLevel2Values = new Set<string>();

    dataset.sortedRecordIds.forEach((id) => {
      const record = dataset.records[id];
      if (record.getFormattedValue(level1) === value1) {
        const value2 = record.getFormattedValue(level2);
        if (value2) {
          uniqueLevel2Values.add(value2);
        }
      }
    });

    // Generate Level 2 Groups
    uniqueLevel2Values.forEach((value2) => {
      const matchingItems = dataset.sortedRecordIds
        .map((id) => dataset.records[id])
        .filter(
          (record) =>
            record.getFormattedValue(level1) === value1 &&
            record.getFormattedValue(level2) === value2
        );

      const itemCount = matchingItems.length;

      matchingItems.forEach((record) => {
        const dynamicItem: DynamicItem = {
          ...Object.keys(record).reduce((acc, key) => {
            acc[key] = record.getFormattedValue(key) || undefined;
            return acc;
          }, {} as DynamicItem),
        };
        items.push(dynamicItem);
      });

      level2Groups.push({
        key: `${value1}-${value2}`,
        name: value2,
        startIndex: startIndex,
        count: itemCount,
        level: 1, // Nested level 2
        isCollapsed: true,
      });

      startIndex += itemCount;
    });

    // Create level1 groups, only adding level2 groups as children
    const itemCount = level2Groups.reduce((sum, group) => sum + group.count, 0);

    groups.push({
      key: value1,
      name: value1,
      startIndex: startIndex - itemCount,
      count: itemCount,
      level: 0, // Top level
      isCollapsed: true,
      children: level2Groups, // Attach level2 groups as children
    });
  });

  return { groups, items };
}
//#endregion
//#region Making columns and items
export function makeColumnAndItems(
  dataset: Dataset,
  level1: string,
  level2: string,
  groups: IGroup[]
): {
  items: DynamicItem[];
  columns: IColumn[];
} {
  const columns: IColumn[] = dataset.columns
    //.slice(2) // Skip the first two columns (level1 and level2) - Can be deleteted if not needed
    .map((column) => ({
      name: column.displayName,
      fieldName: column.name,
      minWidth: column.visualSizeFactor,
      key: column.name,
    }));

  const items: DynamicItem[] = [];
  const groupedRecords: Record<string, Record<string, DynamicItem[]>> = {};

  if (
    !dataset ||
    !Array.isArray(dataset.sortedRecordIds) ||
    !dataset.records ||
    !Array.isArray(dataset.columns)
  ) {
    return { items: [], columns: [] };
  }

  dataset.sortedRecordIds.forEach((id) => {
    const record = dataset.records[id];
    const level1Value = (level1 && record.getFormattedValue(level1)) || "";
    const level2Value = (level2 && record.getFormattedValue(level2)) || "";

    if (!groupedRecords[level1Value]) {
      groupedRecords[level1Value] = {};
    }

    if (!groupedRecords[level1Value][level2Value]) {
      groupedRecords[level1Value][level2Value] = [];
    }

    groupedRecords[level1Value][level2Value].push({
      id: id,
      ...Object.fromEntries(
        dataset.columns.map((column) => [
          column.name,
          record.getValue(column.name) || undefined,
        ])
      ),
    });
  });

  Object.keys(groupedRecords).forEach((level1Key) => {
    const level1Group = groupedRecords[level1Key];

    Object.keys(level1Group).forEach((level2Key) => {
      level1Group[level2Key].forEach((item) => {
        items.push({
          ...item,
          [level1Key]: level1Key,
          [level2Key]: level2Key,
        });
      });
    });
  });

  return { items, columns };
}
//#endregion
//#region Component
export const GroupedListComp = ({
  dataset,
  entityName,
  context,
  currentPage,
  loadPrevPage,
  loadNextPage,
}: IGroupedListProps & {
  context: ComponentFramework.Context<IInputs>;
}): JSX.Element => {
  const [groups, setGroups] = useState<IGroup[]>([]);
  const [items, setItems] = useState<DynamicItem[]>([]);
  const [columns, setColumns] = useState<IColumn[]>([]);
  const [level1, setLevel1] = useState<string>(
    Array.isArray(dataset.columns) && dataset.columns[0]?.name
      ? dataset.columns[0].name
      : ""
  );
  const [level2, setLevel2] = useState<string>(
    Array.isArray(dataset.columns) && dataset.columns[1]?.name
      ? dataset.columns[1].name
      : ""
  );

  const [selectedRecordIds, setSelectedRecordIds] = useState<string[]>([]);

  const selection = useConst(
    () =>
      new Selection({
        onSelectionChanged: () => {
          const selected = selection
            .getSelection()
            .map((item) => ((item as DynamicItem).id as string) ?? "");
          setSelectedRecordIds(selected);
        },
      })
  );

  useEffect(() => {
    if (!level1 || !level2) {
      setGroups([]);
      setItems([]);
      setColumns([]);
      return;
    }
    const { groups: generatedGroups, items: itemsFromMakeGroup } = makeGroups(
      dataset,
      level1,
      level2
    );
    setGroups(generatedGroups);
  }, [dataset, level1, level2]);

  useEffect(() => {
    if (Array.isArray(dataset.columns)) {
      setLevel1(dataset.columns[0]?.name || "");
      setLevel2(dataset.columns[1]?.name || "");
    }
  }, [dataset.columns]);

  useEffect(() => {
    dataset.setSelectedRecordIds(selectedRecordIds);
  }, [selectedRecordIds, dataset]);

  useEffect(() => {
    selection.setItems(items, true);
  }, [items, selection]);

  useEffect(() => {
    if (groups.length > 0) {
      const { items: extractedItems, columns: extractedColumns } =
        makeColumnAndItems(dataset, level1 ?? "", level2 ?? "", groups);
      setItems(extractedItems);
      setColumns(extractedColumns);
    }
  }, [dataset, level1, level2, groups]);

  const openRecordForm = (item: DynamicItem) => {
    context.navigation.openForm({
      entityName: entityName,
      entityId: item.id as string,
    });
  };

  //#region Component render
  return (
    <div className="GroupByCTRL-container" data-control-id="GroupByCTRL">
      {!level1 || !level2 ? (
        <div style={{ padding: 16, color: "red", textAlign: "center" }}>
          Please select at least two columns in the view to enable grouping.
        </div>
      ) : (
        <>
          <DetailsHeader
            columns={columns}
            selection={selection}
            selectionMode={SelectionMode.multiple}
            onColumnClick={() => {}}
            ariaLabelForSelectAllCheckbox="Toggle selection"
            ariaLabelForSelectionColumn="Toggle selection"
            layoutMode={1}
          />
          <SelectionZone
            selection={selection}
            selectionMode={SelectionMode.multiple}
          >
            <GroupedList
              items={items}
              onRenderCell={(
                nestingDepth?: number,
                item?: DynamicItem,
                itemIndex?: number,
                group?: IGroup
              ) =>
                item && typeof itemIndex === "number" && itemIndex > -1 ? (
                  <div onDoubleClick={() => openRecordForm(item)}>
                    <DetailsRow
                      key={item.id as string}
                      columns={columns.map((col) => ({
                        ...col,
                        onRender: (fieldItem: DynamicItem) => {
                          const value = fieldItem[col.fieldName!];
                          if (
                            value &&
                            typeof value === "object" &&
                            ("id" in value || "entityType" in value)
                          ) {
                            return (
                              <span
                                style={{
                                  cursor: "pointer",
                                  textDecoration: "underline",
                                  color: "#0078d4",
                                }}
                                onClick={() => openRecordForm(fieldItem)}
                              >
                                {("name" in value && value.name) ||
                                  ("id" in value && value.id) ||
                                  ""}
                              </span>
                            );
                          }
                          return <span>{value as string}</span>;
                        },
                      }))}
                      groupNestingDepth={nestingDepth}
                      item={item}
                      itemIndex={itemIndex}
                      selection={selection}
                      selectionMode={SelectionMode.multiple}
                      group={group}
                    />
                  </div>
                ) : null
              }
              selection={selection}
              selectionMode={SelectionMode.multiple}
              groups={groups}
              className="ms-GroupedList"
            />
          </SelectionZone>
          {/* Footer */}
          <div className="groupby-footer-bar">
            <div className="footer-info">
              {(() => {
                return `Page: ${currentPage} | Rows showing: ${dataset.sortedRecordIds.length}`;
              })()}
            </div>
            <div className="pagination-buttons">
              <button
                className="pagination-btn"
                onClick={loadPrevPage}
                aria-label="Previous page"
                disabled={!dataset.paging.hasPreviousPage}
              >
                <svg width="20" height="20" viewBox="0 0 20 20" fill="none">
                  <path
                    d="M13 15L8 10L13 5"
                    stroke="#0078d4"
                    strokeWidth="2"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />
                </svg>
              </button>
              <button
                className="pagination-btn"
                onClick={loadNextPage}
                aria-label="Next page"
                disabled={!dataset.paging.hasNextPage}
              >
                <svg width="20" height="20" viewBox="0 0 20 20" fill="none">
                  <path
                    d="M7 5L12 10L7 15"
                    stroke="#0078d4"
                    strokeWidth="2"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  />
                </svg>
              </button>
            </div>
          </div>
        </>
      )}
    </div>
  );
  //#endregion
};
