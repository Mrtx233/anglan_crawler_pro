from __future__ import annotations


class CheckpointStore:
    """Retain failed snapshots until every pending save has succeeded."""

    def __init__(self):
        self.pending = {}

    def save(self, key, writer, *args):
        self.pending[key] = (writer, args)
        writer(*args)
        self.pending.pop(key, None)

    def retry_all(self):
        for key, (writer, args) in list(self.pending.items()):
            self.save(key, writer, *args)
