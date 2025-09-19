import * as React from "react";
import {
    DetailsList, IColumn, IconButton, Stack,
    Text, Rating, IRatingProps, TextField,
    DefaultButton, PrimaryButton, SelectionMode,
    DetailsRow
} from "@fluentui/react";

type Ticket = {
    Id: number;
    Title: string;
    Description?: string;
    Requestor?: { Title?: string; Email?: string } | string;
    AssignedTo?: { Title?: string; Email?: string } | string;
    Status?: string;
    Provider_Rating?: number;
    Comments?: string;
    Manager_description?: string;
};

type Props = {
    tickets: Ticket[];
    onSaveRating?: (ticketId: number, rating: number, comment: string) => Promise<void> | void;
    onSaveManagerDescription?: (ticketId: number, text: string) => Promise<void> | void;
    currentUserEmail?: string;
};

export const CompletedTickets: React.FC<Props> = ({
    tickets,
    onSaveRating,
    onSaveManagerDescription,
    currentUserEmail
}) => {
    const completedTickets = React.useMemo(
        () => tickets.filter((t) => (t.Status || "").toString().toLowerCase() === "completed"),
        [tickets]
    );

    const [ratings, setRatings] = React.useState<Record<number, number>>({});
    const [comments, setComments] = React.useState<Record<number, string>>({});
    const [editingCommentFor, setEditingCommentFor] = React.useState<number | null>(null);
    const [savingFor, setSavingFor] = React.useState<number | null>(null);

    // manager notes states
    const [expandedFor, setExpandedFor] = React.useState<Record<number, boolean>>({});
    const [descriptions, setDescriptions] = React.useState<Record<number, string>>({});
    const [savingDescriptionFor, setSavingDescriptionFor] = React.useState<number | null>(null);


    React.useEffect(() => {
        const rMap: Record<number, number> = {};
        const cMap: Record<number, string> = {};
        const mMap: Record<number, string> = {};

        completedTickets.forEach((t) => {
            const rawR = (t as any).Provider_Rating;
            const parsedR = typeof rawR === "number" ? rawR : Number(rawR) || 0;
            rMap[t.Id] = parsedR;

            const rawC = (t as any).Comments;
            cMap[t.Id] = rawC ? String(rawC) : "";

            const rawM = (t as any).Manager_description;
            mMap[t.Id] = rawM ? String(rawM) : "";
        });

        setRatings(rMap);
        setComments(cMap);
        setDescriptions(mMap);

        // quick debug
        // console.log("CompletedTickets.prefill (first 3):", completedTickets.slice(0, 3).map(x => ({ Id: x.Id, Provider_Rating: (x as any).Provider_Rating, Comments: (x as any).Comments, Manager_description: (x as any).Manager_description }))));
    }, [completedTickets]);


    const onDescriptionChange = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newVal?: string, id?: number) => {
        if (!id) return;
        setDescriptions((prev) => ({ ...prev, [id]: newVal ?? "" }));
    };

    const handleSaveRating = async (id: number) => {
        setSavingFor(id);
        try {
            await onSaveRating?.(id, ratings[id] ?? 0, comments[id] ?? "");
            setEditingCommentFor(null);
        } finally {
            setSavingFor(null);
        }
    };

    const handleSaveDescription = async (id: number) => {
        setSavingDescriptionFor(id);
        try {
            await onSaveManagerDescription?.(id, descriptions[id] ?? "");
            // collapse after save (optional)
            setExpandedFor((prev) => ({ ...prev, [id]: false }));
        } finally {
            setSavingDescriptionFor(null);
        }
    };

    const columns: IColumn[] = [
        {
            key: "colTitle",
            name: "Ticket Title",
            fieldName: "Title",
            minWidth: 100,
            isResizable: true,
            onRender: (item: Ticket) => <Text>{item.Title}</Text>
        },
        {
            key: "colRequestor",
            name: "Seeker",
            fieldName: "Requestor",
            minWidth: 100,
            onRender: (item: Ticket) => {
                const label =
                    typeof item.Requestor === "string"
                        ? item.Requestor
                        : (item.Requestor && (item.Requestor as any).Title) || "";
                return <Text>{label}</Text>;
            }
        },
        {
            key: "colProvider",
            name: "Provider",
            fieldName: "AssignedTo",
            minWidth: 100,
            onRender: (item: Ticket) => {
                const label =
                    typeof item.AssignedTo === "string"
                        ? item.AssignedTo
                        : (item.AssignedTo && (item.AssignedTo as any).Title) || "";
                return <Text>{label}</Text>;
            }
        },
        {
            key: "colStatus",
            name: "Status",
            fieldName: "Status",
            minWidth: 100,
            onRender: (item: Ticket) => <Text>{item.Status}</Text>
        },
        {
            key: "colRate",
            name: "Rate Provider",
            minWidth: 260,
            maxWidth: 380,
            onRender: (item: Ticket) => {
                const id = item.Id;
                const value = ratings[id] ?? 0;

                const onChange: IRatingProps["onChange"] = (_, newValue) => {
                    setRatings((prev) => ({ ...prev, [id]: newValue ?? 0 }));
                    setEditingCommentFor(id);
                };

                const onCommentChangeLocal = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newVal?: string) => {
                    setComments((prev) => ({ ...prev, [id]: newVal ?? "" }));
                };

                return (
                    <Stack tokens={{ childrenGap: 8 }}>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                            <Rating
                                min={1}
                                max={5}
                                rating={value}
                                onChange={onChange}
                                ariaLabelFormat="{0} stars"
                            />
                            <IconButton
                                iconProps={{ iconName: "Comment" }}
                                title="Add comment"
                                ariaLabel="Add comment"
                                onClick={() =>
                                    setEditingCommentFor((prev) => (prev === id ? null : id))
                                }
                            />
                        </Stack>

                        {editingCommentFor === id && (
                            <Stack tokens={{ childrenGap: 8 }}>
                                <TextField
                                    label="Comment (optional)"
                                    multiline
                                    rows={3}
                                    value={comments[id] ?? ""}
                                    onChange={onCommentChangeLocal}
                                />
                                <Stack horizontal tokens={{ childrenGap: 8 }}>
                                    <DefaultButton onClick={() => setEditingCommentFor(null)}>
                                        Cancel
                                    </DefaultButton>
                                    <PrimaryButton
                                        onClick={() => handleSaveRating(id)}
                                        disabled={(ratings[id] ?? 0) === 0 || savingFor === id}
                                    >
                                        {savingFor === id ? "Saving..." : "Save"}
                                    </PrimaryButton>
                                </Stack>
                            </Stack>
                        )}
                    </Stack>
                );
            }
        },
        {
            key: "colDesc",
            name: "Description",
            fieldName: "Manager_description",
            minWidth: 180,
            maxWidth: 220,
            onRender: (item: Ticket) => {
                const id = item.Id;
                const isExpanded = !!expandedFor[id];
                const text = descriptions[id] ?? "";

                return (
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end' }}>
                        {/* Show short preview when collapsed */}
                        {!isExpanded ? (
                            <div style={{ width: '100%', textAlign: 'left' }}>
                                <div style={{ minHeight: 24 }}>
                                    <span style={{ color: '#333' }}>{text ? (text.length > 120 ? text.slice(0, 120) + '…' : text) : <em style={{ color: '#999' }}>No description</em>}</span>
                                </div>
                                <div style={{ marginTop: 8 }}>
                                    <DefaultButton
                                        text="Show Description"
                                        onClick={() => setExpandedFor(prev => ({ ...prev, [id]: true }))}
                                    />
                                </div>
                            </div>
                        ) : (
                            // Expanded editor inside the same cell
                            <div style={{ width: '100%', textAlign: 'left' }}>
                                <TextField
                                    label=""
                                    multiline
                                    rows={3}
                                    value={descriptions[id] ?? ""}
                                    onChange={(e, v) => onDescriptionChange(e, v, id)}
                                    placeholder="Add description..."
                                />
                                <div style={{ marginTop: 8, display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
                                    <DefaultButton onClick={() => setExpandedFor(prev => ({ ...prev, [id]: false }))}>
                                        Cancel
                                    </DefaultButton>
                                    <PrimaryButton
                                        onClick={() => handleSaveDescription(id)}
                                        disabled={savingDescriptionFor === id}
                                    >
                                        {savingDescriptionFor === id ? "Saving..." : "Save"}
                                    </PrimaryButton>
                                </div>
                            </div>
                        )}
                    </div>
                );
            }
        }

    ];

    // Add a custom "Manager notes" row renderer below each item — we will render this as a DetailsList custom row by adding an extra blank column to allow us to display the expanded panel.
    // Simpler approach: keep the DetailsList as-is and render the manager note editor under the list (by mapping items). That's easier to control layout.

    return (
        <Stack tokens={{ childrenGap: 12 }}>
            <Stack horizontal verticalAlign="center" horizontalAlign="space-between">
                <Text variant="large">Completed Requests</Text>
                <Text>{completedTickets.length} item(s)</Text>
            </Stack>

            {completedTickets.length === 0 ? (
                <Text>No completed tickets to rate.</Text>
            ) : (
                // We'll render a simple table-like list using DetailsList and then map expanded manager notes below each row.
                <div>
                    <DetailsList
                        items={completedTickets}
                        columns={columns}
                        setKey="completedList"
                        selectionMode={SelectionMode.none}
onRenderRow={(props) => {
  if (!props) return null;
  // Render standard DetailsRow only. Description editor lives in the Description column.
  return <DetailsRow {...props} />;
}}

                    />
                </div>
            )}
        </Stack>
    );
};

export default CompletedTickets;
