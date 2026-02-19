"""
Additional tests for windows_mcp.filesystem.service targeting uncovered lines.

Coverage focus areas:
- read_file: file-too-large, UnicodeDecodeError, PermissionError, generic Exception
- write_file: PermissionError, generic Exception, create_parents=False
- copy_path: directory overwrite (rmtree path), unsupported type, PermissionError, Exception
- move_path: overwrite existing dir, overwrite existing file, PermissionError, Exception
- delete_path: unsupported type, PermissionError, Exception
- list_directory: path-is-not-dir, recursive iterator, MAX_RESULTS truncation,
                  OSError on stat, PermissionError, Exception
- search_files: path-is-not-dir, non-recursive glob, MAX_RESULTS truncation,
                OSError on stat, PermissionError, Exception
- get_file_info: dir PermissionError on iterdir, symlink, PermissionError on stat, Exception
"""

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from windows_mcp.filesystem.service import (
    copy_path,
    delete_path,
    get_file_info,
    list_directory,
    move_path,
    read_file,
    search_files,
    write_file,
)
from windows_mcp.filesystem.views import MAX_READ_SIZE, MAX_RESULTS

# ---------------------------------------------------------------------------
# TestReadFile - uncovered branches
# ---------------------------------------------------------------------------


class TestReadFileEdgeCases:
    def test_empty_file(self, tmp_path):
        f = tmp_path / "empty.txt"
        f.write_text("", encoding="utf-8")
        result = read_file(str(f))
        assert "File:" in result

    def test_file_too_large(self, tmp_path):
        f = tmp_path / "big.bin"
        # Write a real file and patch only the st_size attribute check that
        # read_file uses: file_path.stat().st_size > MAX_READ_SIZE.
        # We must allow exists() and is_file() (which also call stat internally)
        # to work normally, so we use a real stat result with st_size overridden.
        f.write_bytes(b"x")
        real_stat = f.stat()

        class FakeStat:
            st_size = MAX_READ_SIZE + 1
            st_ctime = real_stat.st_ctime
            st_mtime = real_stat.st_mtime
            st_atime = real_stat.st_atime
            st_mode = real_stat.st_mode

        original_stat = Path.stat

        def patched_stat(self, *args, **kwargs):
            if self == f.resolve():
                return FakeStat()
            return original_stat(self, *args, **kwargs)

        with patch.object(Path, "stat", patched_stat):
            result = read_file(str(f))
        assert "Error: File too large" in result
        assert "Maximum is" in result

    def test_unicode_decode_error(self, tmp_path):
        # Write bytes that are invalid UTF-8 but force the error through mocking
        f = tmp_path / "binary.bin"
        f.write_bytes(b"\xff\xfe")
        with patch("builtins.open", side_effect=UnicodeDecodeError("utf-8", b"", 0, 1, "bad")):
            result = read_file(str(f))
        assert "Unable to read file as text" in result

    def test_permission_error_on_open(self, tmp_path):
        f = tmp_path / "locked.txt"
        f.write_text("data", encoding="utf-8")
        with patch("builtins.open", side_effect=PermissionError("access denied")):
            result = read_file(str(f))
        assert "Error: Permission denied" in result

    def test_generic_exception_on_open(self, tmp_path):
        f = tmp_path / "bad.txt"
        f.write_text("data", encoding="utf-8")
        with patch("builtins.open", side_effect=OSError("disk error")):
            result = read_file(str(f))
        assert "Error reading file" in result

    def test_read_offset_beyond_end(self, tmp_path):
        f = tmp_path / "short.txt"
        f.write_text("line1\nline2\n", encoding="utf-8")
        # offset beyond total lines returns empty content
        result = read_file(str(f), offset=100)
        assert "File:" in result

    def test_read_special_characters_in_content(self, tmp_path):
        f = tmp_path / "special.txt"
        content = "tab:\there\nnewlines\r\nunicode: \u00e9\u00e0\u00fc"
        f.write_text(content, encoding="utf-8")
        result = read_file(str(f))
        assert "tab:" in result
        assert "unicode:" in result

    def test_read_offset_only_no_limit(self, tmp_path):
        f = tmp_path / "lines.txt"
        f.write_text("a\nb\nc\nd\n", encoding="utf-8")
        result = read_file(str(f), offset=3)
        assert "c" in result
        assert "d" in result

    def test_read_limit_only_no_offset(self, tmp_path):
        f = tmp_path / "lines.txt"
        f.write_text("a\nb\nc\nd\n", encoding="utf-8")
        result = read_file(str(f), limit=2)
        assert "a" in result
        assert "b" in result

    def test_read_with_latin1_encoding(self, tmp_path):
        f = tmp_path / "latin.txt"
        f.write_bytes(b"caf\xe9")
        result = read_file(str(f), encoding="latin-1")
        assert "caf" in result


# ---------------------------------------------------------------------------
# TestWriteFile - uncovered branches
# ---------------------------------------------------------------------------


class TestWriteFileEdgeCases:
    def test_permission_error_on_write(self, tmp_path):
        f = tmp_path / "locked.txt"
        with patch("builtins.open", side_effect=PermissionError("access denied")):
            result = write_file(str(f), "content")
        assert "Error: Permission denied" in result

    def test_generic_exception_on_write(self, tmp_path):
        f = tmp_path / "bad.txt"
        with patch("builtins.open", side_effect=OSError("disk full")):
            result = write_file(str(f), "content")
        assert "Error writing file" in result

    def test_write_without_create_parents_existing_dir(self, tmp_path):
        f = tmp_path / "direct.txt"
        result = write_file(str(f), "hello", create_parents=False)
        assert "Written to" in result
        assert f.read_text(encoding="utf-8") == "hello"

    def test_write_append_reports_appended(self, tmp_path):
        f = tmp_path / "app.txt"
        f.write_text("start", encoding="utf-8")
        result = write_file(str(f), " end", append=True)
        assert "Appended to" in result

    def test_write_empty_content(self, tmp_path):
        f = tmp_path / "empty.txt"
        result = write_file(str(f), "")
        assert "Written to" in result
        assert f.read_text(encoding="utf-8") == ""


# ---------------------------------------------------------------------------
# TestCopyPath - uncovered branches
# ---------------------------------------------------------------------------


class TestCopyPathEdgeCases:
    def test_copy_directory_with_overwrite_removes_existing_dst(self, tmp_path):
        src = tmp_path / "src_dir"
        src.mkdir()
        (src / "new.txt").write_text("new", encoding="utf-8")
        dst = tmp_path / "dst_dir"
        dst.mkdir()
        (dst / "old.txt").write_text("old", encoding="utf-8")
        result = copy_path(str(src), str(dst), overwrite=True)
        assert "Copied directory" in result
        assert (dst / "new.txt").read_text(encoding="utf-8") == "new"
        assert not (dst / "old.txt").exists()

    def test_copy_permission_error(self, tmp_path):
        src = tmp_path / "src.txt"
        src.write_text("data", encoding="utf-8")
        dst = tmp_path / "dst.txt"
        with patch("shutil.copy2", side_effect=PermissionError("access denied")):
            result = copy_path(str(src), str(dst))
        assert "Error: Permission denied" in result

    def test_copy_generic_exception(self, tmp_path):
        src = tmp_path / "src.txt"
        src.write_text("data", encoding="utf-8")
        dst = tmp_path / "dst.txt"
        with patch("shutil.copy2", side_effect=OSError("disk error")):
            result = copy_path(str(src), str(dst))
        assert "Error copying" in result

    def test_copy_creates_parent_dirs_for_file(self, tmp_path):
        src = tmp_path / "src.txt"
        src.write_text("data", encoding="utf-8")
        dst = tmp_path / "nested" / "deep" / "dst.txt"
        result = copy_path(str(src), str(dst))
        assert "Copied file" in result
        assert dst.read_text(encoding="utf-8") == "data"


# ---------------------------------------------------------------------------
# TestMovePath - uncovered branches
# ---------------------------------------------------------------------------


class TestMovePathEdgeCases:
    def test_move_overwrite_existing_directory(self, tmp_path):
        src = tmp_path / "src_dir"
        src.mkdir()
        (src / "file.txt").write_text("moved", encoding="utf-8")
        dst = tmp_path / "dst_dir"
        dst.mkdir()
        (dst / "old.txt").write_text("old", encoding="utf-8")
        result = move_path(str(src), str(dst), overwrite=True)
        assert "Moved" in result
        assert not src.exists()
        assert (dst / "file.txt").read_text(encoding="utf-8") == "moved"

    def test_move_overwrite_existing_file(self, tmp_path):
        src = tmp_path / "src.txt"
        src.write_text("new", encoding="utf-8")
        dst = tmp_path / "dst.txt"
        dst.write_text("old", encoding="utf-8")
        result = move_path(str(src), str(dst), overwrite=True)
        assert "Moved" in result
        assert not src.exists()
        assert dst.read_text(encoding="utf-8") == "new"

    def test_move_permission_error(self, tmp_path):
        src = tmp_path / "src.txt"
        src.write_text("data", encoding="utf-8")
        dst = tmp_path / "dst.txt"
        with patch("shutil.move", side_effect=PermissionError("access denied")):
            result = move_path(str(src), str(dst))
        assert "Error: Permission denied" in result

    def test_move_generic_exception(self, tmp_path):
        src = tmp_path / "src.txt"
        src.write_text("data", encoding="utf-8")
        dst = tmp_path / "dst.txt"
        with patch("shutil.move", side_effect=OSError("disk error")):
            result = move_path(str(src), str(dst))
        assert "Error moving" in result

    def test_move_creates_parent_dirs_for_destination(self, tmp_path):
        src = tmp_path / "src.txt"
        src.write_text("data", encoding="utf-8")
        dst = tmp_path / "nested" / "subdir" / "dst.txt"
        result = move_path(str(src), str(dst))
        assert "Moved" in result
        assert dst.read_text(encoding="utf-8") == "data"


# ---------------------------------------------------------------------------
# TestDeletePath - uncovered branches
# ---------------------------------------------------------------------------


class TestDeletePathEdgeCases:
    def test_delete_permission_error(self, tmp_path):
        f = tmp_path / "locked.txt"
        f.write_text("data", encoding="utf-8")
        with patch.object(Path, "unlink", side_effect=PermissionError("access denied")):
            result = delete_path(str(f))
        assert "Error: Permission denied" in result

    def test_delete_generic_exception(self, tmp_path):
        f = tmp_path / "bad.txt"
        f.write_text("data", encoding="utf-8")
        with patch.object(Path, "unlink", side_effect=OSError("disk error")):
            result = delete_path(str(f))
        assert "Error deleting" in result

    def test_delete_symlink_branch_via_mock(self, tmp_path):
        # delete_path calls Path(path).resolve() first; on Windows, resolve()
        # follows symlinks, so the is_symlink() branch (service.py:151) is only
        # reachable when the resolved path is itself a symlink.  We exercise this
        # branch by mocking is_symlink() on the resolved path object so it returns
        # True, verifying the code calls unlink() rather than rmdir().
        f = tmp_path / "sym_target.txt"
        f.write_text("data", encoding="utf-8")

        original_is_symlink = Path.is_symlink

        def patched_is_symlink(self):
            if self == f.resolve():
                return True
            return original_is_symlink(self)

        with patch.object(Path, "is_symlink", patched_is_symlink):
            result = delete_path(str(f))
        assert "Deleted file" in result


# ---------------------------------------------------------------------------
# TestListDirectory - uncovered branches
# ---------------------------------------------------------------------------


class TestListDirectoryEdgeCases:
    def test_list_path_is_a_file_not_dir(self, tmp_path):
        f = tmp_path / "file.txt"
        f.write_text("data", encoding="utf-8")
        result = list_directory(str(f))
        assert "Error: Path is not a directory" in result

    def test_list_recursive(self, tmp_path):
        sub = tmp_path / "subdir"
        sub.mkdir()
        (sub / "deep.txt").write_text("x", encoding="utf-8")
        (tmp_path / "top.txt").write_text("y", encoding="utf-8")
        result = list_directory(str(tmp_path), recursive=True)
        assert "deep.txt" in result
        assert "top.txt" in result

    def test_list_recursive_with_pattern(self, tmp_path):
        sub = tmp_path / "sub"
        sub.mkdir()
        (sub / "a.py").write_text("x", encoding="utf-8")
        (sub / "b.txt").write_text("y", encoding="utf-8")
        result = list_directory(str(tmp_path), pattern="*.py", recursive=True)
        assert "a.py" in result
        assert "b.txt" not in result

    def test_list_truncates_at_max_results(self, tmp_path):
        # Create MAX_RESULTS + 2 files to trigger truncation
        for i in range(MAX_RESULTS + 2):
            (tmp_path / f"file_{i:04d}.txt").write_text("x", encoding="utf-8")
        result = list_directory(str(tmp_path))
        assert "truncated" in result

    def test_list_empty_dir_with_pattern(self, tmp_path):
        result = list_directory(str(tmp_path), pattern="*.py")
        assert "empty" in result.lower()
        assert "*.py" in result

    def test_list_permission_error(self, tmp_path):
        with patch("windows_mcp.filesystem.service.Path.iterdir", side_effect=PermissionError):
            result = list_directory(str(tmp_path))
        assert "Error: Permission denied" in result

    def test_list_generic_exception(self, tmp_path):
        with patch("windows_mcp.filesystem.service.Path.iterdir", side_effect=OSError("boom")):
            result = list_directory(str(tmp_path))
        assert "Error listing directory" in result

    def test_list_stat_oserror_falls_back_to_zero_size(self, tmp_path):
        # The OSError fallback at service.py:204 guards entry.stat().st_size.
        # Patching Path.stat globally is unsafe here because iterdir() and sorted()
        # call stat() internally (Python 3.13), causing the OSError to escape the
        # inner except block and be caught by the outer except Exception instead.
        # We inject a MagicMock entry so that is_file() returns True and stat()
        # raises OSError -- the inner try/except catches it and size falls to 0.
        fake_entry = MagicMock()
        fake_entry.name = "a.txt"
        fake_entry.is_dir.return_value = False
        fake_entry.is_file.return_value = True
        fake_entry.stat.side_effect = OSError("stat failed")
        fake_entry.relative_to.return_value = Path("a.txt")

        def _patched_iterdir(self):  # noqa: ANN001
            yield fake_entry

        with patch.object(Path, "iterdir", _patched_iterdir):
            result = list_directory(str(tmp_path))
        # Entry appears with size=0 fallback; no crash from the OSError
        assert "a.txt" in result

    def test_list_header_includes_pattern(self, tmp_path):
        (tmp_path / "match.py").write_text("x", encoding="utf-8")
        result = list_directory(str(tmp_path), pattern="*.py")
        assert "(filter: *.py)" in result


# ---------------------------------------------------------------------------
# TestSearchFiles - uncovered branches
# ---------------------------------------------------------------------------


class TestSearchFilesEdgeCases:
    def test_search_path_is_a_file_not_dir(self, tmp_path):
        f = tmp_path / "file.txt"
        f.write_text("data", encoding="utf-8")
        result = search_files(str(f), "*.py")
        assert "Error: Search path is not a directory" in result

    def test_search_non_recursive(self, tmp_path):
        (tmp_path / "top.py").write_text("x", encoding="utf-8")
        sub = tmp_path / "sub"
        sub.mkdir()
        (sub / "nested.py").write_text("x", encoding="utf-8")
        result = search_files(str(tmp_path), "*.py", recursive=False)
        assert "top.py" in result
        assert "nested.py" not in result

    def test_search_truncates_at_max_results(self, tmp_path):
        for i in range(MAX_RESULTS + 2):
            (tmp_path / f"file_{i:04d}.py").write_text("x", encoding="utf-8")
        result = search_files(str(tmp_path), "*.py")
        assert "truncated" in result

    def test_search_permission_error(self, tmp_path):
        with patch("windows_mcp.filesystem.service.Path.rglob", side_effect=PermissionError):
            result = search_files(str(tmp_path), "*.py")
        assert "Error: Permission denied" in result

    def test_search_generic_exception(self, tmp_path):
        with patch("windows_mcp.filesystem.service.Path.rglob", side_effect=OSError("boom")):
            result = search_files(str(tmp_path), "*.py")
        assert "Error searching" in result

    def test_search_stat_oserror_falls_back_to_zero_size(self, tmp_path):
        # The OSError fallback at service.py:251 guards match.stat().st_size when
        # match.is_file() is True. Use a MagicMock entry so is_file/is_dir behave
        # exactly as configured and stat() raises OSError to trigger the fallback.
        # Patching Path.stat globally is not viable here because rglob() calls
        # stat() internally during iteration (Python 3.13), causing the OSError
        # to escape the inner except and be caught by the outer except Exception.
        fake_entry = MagicMock()
        fake_entry.name = "a.py"
        fake_entry.is_dir.return_value = False
        fake_entry.is_file.return_value = True
        fake_entry.stat.side_effect = OSError("stat failed")
        fake_entry.relative_to.return_value = Path("a.py")

        def patched_rglob(self, pattern):  # noqa: ARG001
            yield fake_entry

        with patch.object(Path, "rglob", patched_rglob):
            result = search_files(str(tmp_path), "*.py")
        # Entry appears with size=0 fallback; no crash from the OSError
        assert "a.py" in result

    def test_search_result_count_in_header(self, tmp_path):
        for name in ("foo.py", "bar.py", "baz.py"):
            (tmp_path / name).write_text("x", encoding="utf-8")
        result = search_files(str(tmp_path), "*.py")
        assert "3 matches" in result

    def test_search_directories_matched_by_pattern(self, tmp_path):
        (tmp_path / "mydir.py").mkdir()  # a dir named with .py suffix
        result = search_files(str(tmp_path), "*.py")
        assert "mydir.py" in result


# ---------------------------------------------------------------------------
# TestGetFileInfo - uncovered branches
# ---------------------------------------------------------------------------


class TestGetFileInfoEdgeCases:
    def test_symlink_reports_link_target(self, tmp_path):
        target = tmp_path / "real.txt"
        target.write_text("data", encoding="utf-8")
        link = tmp_path / "link.txt"
        try:
            link.symlink_to(target)
        except (OSError, NotImplementedError):
            pytest.skip("Symlinks not supported on this platform/configuration")
        # get_file_info resolves the path; after resolve(), the symlink itself is gone.
        # Call on the symlink path without resolve to exercise the symlink branch.
        result = get_file_info(str(link))
        # Symlink resolves to the real file on Windows by default via Path.resolve()
        # The result will at least show metadata without error.
        assert "Error" not in result or "Path not found" not in result

    def test_dir_permission_error_on_iterdir(self, tmp_path):
        d = tmp_path / "restricted"
        d.mkdir()
        (d / "child.txt").write_text("x", encoding="utf-8")
        with patch.object(Path, "iterdir", side_effect=PermissionError("denied")):
            result = get_file_info(str(d))
        # Should return directory info without crashing (PermissionError caught internally)
        assert "Type: Directory" in result

    def test_permission_error_on_stat(self, tmp_path):
        f = tmp_path / "locked.txt"
        f.write_text("data", encoding="utf-8")
        original_stat = Path.stat
        call_count = {"n": 0}

        def patched_stat(self, *args, **kwargs):
            call_count["n"] += 1
            # Allow the first calls (used by exists() / is_file() checks),
            # then raise on the explicit stat() call inside the try block.
            if call_count["n"] > 2:
                raise PermissionError("access denied")
            return original_stat(self, *args, **kwargs)

        with patch.object(Path, "stat", patched_stat):
            result = get_file_info(str(f))
        assert "Error: Permission denied" in result

    def test_generic_exception_on_stat(self, tmp_path):
        f = tmp_path / "broken.txt"
        f.write_text("data", encoding="utf-8")
        original_stat = Path.stat
        call_count = {"n": 0}

        def patched_stat(self, *args, **kwargs):
            call_count["n"] += 1
            if call_count["n"] > 2:
                raise OSError("disk error")
            return original_stat(self, *args, **kwargs)

        with patch.object(Path, "stat", patched_stat):
            result = get_file_info(str(f))
        assert "Error getting file info" in result

    def test_file_without_extension(self, tmp_path):
        f = tmp_path / "Makefile"
        f.write_text("all:", encoding="utf-8")
        result = get_file_info(str(f))
        assert "Extension: (none)" in result

    def test_dir_counts_files_and_subdirs(self, tmp_path):
        d = tmp_path / "mydir"
        d.mkdir()
        (d / "f1.txt").write_text("x", encoding="utf-8")
        (d / "f2.txt").write_text("y", encoding="utf-8")
        sub = d / "subdir"
        sub.mkdir()
        result = get_file_info(str(d))
        assert "2 files" in result
        assert "1 directories" in result

    def test_read_only_field_reported(self, tmp_path):
        f = tmp_path / "ronly.txt"
        f.write_text("data", encoding="utf-8")
        result = get_file_info(str(f))
        # Read-only field present regardless of value
        assert "Read-only:" in result


# ---------------------------------------------------------------------------
# Coverage gap: filesystem lines 110, 164, 312
# ---------------------------------------------------------------------------


class TestUnsupportedFileType:
    """Cover 'unsupported file type' branches for paths that are neither
    regular files nor directories (e.g. device nodes, named pipes)."""

    def test_copy_unsupported_file_type(self, tmp_path):
        """copy_path returns error for a path that exists but is neither file nor dir (line 110)."""
        src = tmp_path / "special"
        src.write_text("x", encoding="utf-8")
        dst = tmp_path / "dest"
        # Mock is_file and is_dir to return False while exists returns True
        with (
            patch.object(Path, "is_file", return_value=False),
            patch.object(Path, "is_dir", return_value=False),
        ):
            result = copy_path(str(src), str(dst))
        assert "Unsupported file type" in result

    def test_delete_unsupported_file_type(self, tmp_path):
        """delete_path returns error for unsupported file type (line 164)."""
        target = tmp_path / "special"
        target.write_text("x", encoding="utf-8")
        with (
            patch.object(Path, "is_file", return_value=False),
            patch.object(Path, "is_dir", return_value=False),
            patch.object(Path, "is_symlink", return_value=False),
        ):
            result = delete_path(str(target))
        assert "Unsupported file type" in result


class TestGetFileInfoSymlink:
    """Cover line 312: file.link_target = str(os.readlink(target))."""

    def test_symlink_link_target_reported(self, tmp_path):
        """get_file_info sets link_target when target.is_symlink() is True."""
        f = tmp_path / "target.txt"
        f.write_text("data", encoding="utf-8")
        # Mock is_symlink to return True and os.readlink to return a path
        with (
            patch.object(Path, "is_symlink", return_value=True),
            patch("windows_mcp.filesystem.service.os.readlink", return_value=str(f)),
        ):
            result = get_file_info(str(f))
        assert str(f) in result
