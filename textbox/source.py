import contextlib
import hashlib
import io
import os


class RandomAccessStream(io.RawIOBase):
    """Expose a forensic VFS read_random(offset, size) object as a file."""

    def __init__(self, virtual_file):
        self.virtual_file = virtual_file
        self.position = 0
        size = virtual_file.size() if callable(virtual_file.size) else virtual_file.size
        self.length = int(size)

    def readable(self):
        return True

    def seekable(self):
        return True

    def tell(self):
        return self.position

    def seek(self, offset, whence=io.SEEK_SET):
        if whence == io.SEEK_SET:
            position = offset
        elif whence == io.SEEK_CUR:
            position = self.position + offset
        elif whence == io.SEEK_END:
            position = self.length + offset
        else:
            raise ValueError("invalid whence")
        if position < 0:
            raise ValueError("negative seek position")
        self.position = min(position, self.length)
        return self.position

    def readinto(self, buffer):
        if self.position >= self.length:
            return 0
        size = min(len(buffer), self.length - self.position)
        data = self.virtual_file.read_random(self.position, size)
        if data is None:
            return 0
        data = bytes(data)
        buffer[: len(data)] = data
        self.position += len(data)
        return len(data)


class DocumentSource(object):
    """Normalize paths, bytes, streams, and forensic virtual files."""

    def __init__(self, source, name=None):
        if isinstance(source, DocumentSource):
            self._kind = source._kind
            self._value = source._value
            self.name = name or source.name
            self.size = source.size
            self._materialized = source._materialized
            self._sha256 = source._sha256
            return

        self._materialized = None
        self._sha256 = None
        if isinstance(source, (str, os.PathLike)):
            path = os.fspath(source)
            self._kind = "path"
            self._value = path
            self.name = name or os.path.basename(path)
            self.size = os.path.getsize(path)
        elif isinstance(source, (bytes, bytearray, memoryview)):
            data = bytes(source)
            self._kind = "bytes"
            self._value = data
            self.name = name or "virtual-file"
            self.size = len(data)
        elif hasattr(source, "read_random") and hasattr(source, "size"):
            self._kind = "random-access"
            self._value = source
            self.name = name or self._infer_name(source)
            size = source.size() if callable(source.size) else source.size
            self.size = int(size)
        elif hasattr(source, "open") and not hasattr(source, "read"):
            self._kind = "opener"
            self._value = source
            self.name = name or self._infer_name(source)
            size = getattr(source, "size", None)
            self.size = size() if callable(size) else size
        elif hasattr(source, "read"):
            self._kind = "stream"
            self._value = source
            self.name = name or self._infer_name(source)
            self.size = self._stream_size(source)
        else:
            raise TypeError(
                "source must be a path, bytes, binary stream, or virtual file"
            )

    @staticmethod
    def _infer_name(source):
        name = getattr(source, "name", None)
        return os.path.basename(os.fspath(name)) if name else "virtual-file"

    @staticmethod
    def _stream_size(stream):
        if not hasattr(stream, "seek") or not hasattr(stream, "tell"):
            return None
        try:
            position = stream.tell()
            stream.seek(0, io.SEEK_END)
            size = stream.tell()
            stream.seek(position)
            return size
        except (OSError, ValueError, io.UnsupportedOperation):
            return None

    @contextlib.contextmanager
    def open(self):
        if self._kind == "path":
            with open(self._value, "rb") as stream:
                yield stream
            return
        if self._kind == "bytes":
            with io.BytesIO(self._value) as stream:
                yield stream
            return
        if self._kind == "random-access":
            with io.BufferedReader(RandomAccessStream(self._value)) as stream:
                yield stream
            return
        if self._kind == "opener":
            opened = self._value.open()
            if hasattr(opened, "__enter__"):
                with opened as stream:
                    yield self._ensure_seekable(stream)
            else:
                try:
                    yield self._ensure_seekable(opened)
                finally:
                    close = getattr(opened, "close", None)
                    if close:
                        close()
            return

        # A caller-owned stream cannot safely provide multiple independent
        # readers. Materialize it once and return a fresh BytesIO each time.
        if self._materialized is None:
            stream = self._value
            original_position = None
            try:
                if hasattr(stream, "tell"):
                    original_position = stream.tell()
                if hasattr(stream, "seek"):
                    stream.seek(0)
                self._materialized = bytes(stream.read())
                self.size = len(self._materialized)
            finally:
                if original_position is not None and hasattr(stream, "seek"):
                    stream.seek(original_position)
        with io.BytesIO(self._materialized) as stream:
            yield stream

    def _ensure_seekable(self, stream):
        if hasattr(stream, "seekable") and stream.seekable():
            return stream
        return io.BytesIO(stream.read())

    def read_bytes(self):
        with self.open() as stream:
            return stream.read()

    def read_prefix(self, size=4096):
        with self.open() as stream:
            return stream.read(size)

    def sha256(self):
        if self._sha256 is None:
            digest = hashlib.sha256()
            with self.open() as stream:
                while True:
                    chunk = stream.read(1024 * 1024)
                    if not chunk:
                        break
                    digest.update(chunk)
            self._sha256 = digest.hexdigest()
        return self._sha256

    def provenance(self, detected_format=None):
        return {
            "name": self.name,
            "size": self.size,
            "sha256": self.sha256(),
            "detectedFormat": detected_format,
        }


def as_source(source, name=None):
    return source if isinstance(source, DocumentSource) and name is None else DocumentSource(source, name)
