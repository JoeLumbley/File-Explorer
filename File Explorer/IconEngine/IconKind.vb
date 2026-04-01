'' IconEngine small types and enums.
'' These are intentionally simple so they’re easy to reason about and reuse.

'' How an icon is keyed in the cache.
'Public Enum IconKind
'    ' Icons keyed by file extension (e.g., ".txt").
'    FileExtension

'    ' Icons keyed by full path (e.g., EXEs, custom icons).
'    Path

'    ' Icons keyed by virtual folder canonical name.
'    Virtual

'    ' Icons keyed by folder type (e.g., Pictures, Music).
'    FolderType
'End Enum

'' Where an icon came from. Useful for debugging and tuning.
'Public Enum IconSourceKind
'    Cache
'    Shell
'    Thumbnail
'    Placeholder
'    Fallback
'End Enum

'' Optional overlays that can be composited on top of base icons.
'Public Enum IconOverlayKind
'    None
'    Shortcut
'    SharedFolder
'End Enum