import * as React from "react";
import {
    DetailsList,
    IColumn,
    IconButton,
    Stack,
    Text,
    Rating,
    IRatingProps,
    TextField,
    DefaultButton,
    PrimaryButton,
    SelectionMode
} from "@fluentui/react";

type Ticket = {
    Id: number;
    Title: string;
    Requestor?: { Title?: string; Email?: string } | string;
    AssignedTo?: { Title?: string; Email?: string } | string;
    Status?: string;
    Provider_Rating?: number;
    Comments?: string;
};

type Props = {
    tickets: Ticket[];
    onSaveRating?: (ticketId: number, rating: number, comment: string) => Promise<void> | void;
    currentUserEmail?: string;
};

export const CompletedTickets: React.FC<Props> = ({ tickets, onSaveRating, currentUserEmail }) => {
    const completedTickets = React.useMemo(
        () => tickets.filter((t) => (t.Status || "").toString().toLowerCase() === "completed"),
        [tickets]
    );

    const [ratings, setRatings] = React.useState<Record<number, number>>({});
    const [comments, setComments] = React.useState<Record<number, string>>({});
    const [editingCommentFor, setEditingCommentFor] = React.useState<number | null>(null);
    const [savingFor, setSavingFor] = React.useState<number | null>(null);

    React.useEffect(() => {
        const rMap: Record<number, number> = {};
        const cMap: Record<number, string> = {};

        completedTickets.forEach((t) => {
            const rawR = (t as any).Provider_Rating;
            const parsedR = typeof rawR === "number" ? rawR : Number(rawR) || 0;
            rMap[t.Id] = parsedR;

            const rawC = (t as any).Comments;
            cMap[t.Id] = rawC ? String(rawC) : "";
        });

        setRatings(rMap);
        setComments(cMap);

        // quick debug: verify fields exist
        console.log("CompletedTickets.prefill (first 3):", completedTickets.slice(0, 3).map(x => ({
            Id: x.Id,
            Provider_Rating: (x as any).Provider_Rating,
            Comments: (x as any).Comments
        })));
    }, [completedTickets]);

    const columns: IColumn[] = [
        {
            key: "colTitle",
            name: "Ticket Title",
            fieldName: "Title",
            minWidth: 120,
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
            minWidth: 200,
            maxWidth: 320,
            onRender: (item: Ticket) => {
                const id = item.Id;
                const value = ratings[id] ?? 0;

                const onChange: IRatingProps["onChange"] = (_, newValue) => {
                    setRatings((prev) => ({ ...prev, [id]: newValue ?? 0 }));
                    setEditingCommentFor(id);
                };

                const onCommentChange = (
                    ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
                    newVal?: string
                ) => {
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
                                    onChange={onCommentChange}
                                />
                                <Stack horizontal tokens={{ childrenGap: 8 }}>
                                    <DefaultButton onClick={() => setEditingCommentFor(null)}>
                                        Cancel
                                    </DefaultButton>
<PrimaryButton
  onClick={async () => {
    setSavingFor(id);
    try {
      await onSaveRating?.(id, ratings[id] ?? 0, comments[id] ?? '');
      setEditingCommentFor(null);
    } finally {
      setSavingFor(null);
    }
  }}
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
        }
    ];

    return (
        <Stack tokens={{ childrenGap: 12 }}>
            <Stack horizontal verticalAlign="center" horizontalAlign="space-between">
                <Text variant="large">Completed Requests</Text>
                <Text>{completedTickets.length} item(s)</Text>
            </Stack>

            {completedTickets.length === 0 ? (
                <Text>No completed tickets to rate.</Text>
            ) : (
                <DetailsList
                    items={completedTickets}
                    columns={columns}
                    setKey="completedList"
                    selectionMode={SelectionMode.none}
                />
            )}
        </Stack>
    );
};

export default CompletedTickets;
