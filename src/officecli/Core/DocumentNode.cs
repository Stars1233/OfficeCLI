// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Represents a node in the document DOM tree.
/// This is the universal abstraction across Word/Excel/PowerPoint.
/// </summary>
public class DocumentNode
{
    public string Path { get; set; } = "";
    public string Type { get; set; } = "";
    public string? Text { get; set; }
    public string? Preview { get; set; }
    public string? Style { get; set; }
    public int ChildCount { get; set; }
    public Dictionary<string, object?> Format { get; set; } = new();
    public List<DocumentNode> Children { get; set; } = new();
}
