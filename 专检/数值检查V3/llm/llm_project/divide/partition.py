def smart_split_with_buffer(src_path, num_parts, output_dir, lang_type, buffer_chars=2000,
                            split_element_ratios=None):
    """智能分割文档：按字数均分 + 缓冲区重叠

    Args:
        split_element_ratios: 可选，由主文档计算出的分割比例列表（元素位置占比）。
                              提供时，本文档按相同比例分割，确保原文/译文内容对齐。
    Returns:
        (generated_files, part_info, element_ratios)
        element_ratios: 理想分割点的元素位置比例，可传给另一文档以保持同步。
    """
    doc = Document(src_path)
    elements = get_all_content_elements(doc)
    base_name = os.path.splitext(os.path.basename(src_path))[0]

    # 计算每个元素的累计字数
    element_counts = []
    cumulative_count = 0
    for elem in elements:
        count = get_element_text_count(elem, lang_type)
        cumulative_count += count
        element_counts.append(cumulative_count)

    total_count = cumulative_count
    if total_count == 0:
        log_manager.log_exception("文档字数为0，无法分割")
        return [], [], []

    target_per_part = total_count // num_parts

    log_manager.log(f"总字数: {total_count:,}, 目标每份: {target_per_part:,}, 缓冲区: {buffer_chars:,} 字")

    # 计算理想分割点
    if split_element_ratios is not None:
        # 从主文档的元素比例映射到本文档的元素索引
        log_manager.log(f"使用主文档分割比例: {[f'{r:.4f}' for r in split_element_ratios]}")
        ideal_splits = []
        for ratio in split_element_ratios:
            idx = max(0, min(int(ratio * len(elements)), len(elements) - 1))
            ideal_splits.append(idx)
    else:
        # 自主计算（作为主文档）
        ideal_splits = []
        for i in range(1, num_parts):
            target_chars = target_per_part * i
            split_idx = find_element_index_by_char_count(element_counts, target_chars)
            ideal_splits.append(split_idx)

    # 确保分割点严格递增（避免大元素导致多个分割点重叠）
    for i in range(1, len(ideal_splits)):
        if ideal_splits[i] <= ideal_splits[i - 1]:
            ideal_splits[i] = ideal_splits[i - 1] + 1
    # 确保不越界
    for i in range(len(ideal_splits)):
        ideal_splits[i] = min(ideal_splits[i], len(elements) - 1)

    # 计算元素位置比例（供另一文档使用）
    element_ratios = [idx / len(elements) for idx in ideal_splits] if elements else []

    log_manager.log(f"理想分割点索引: {ideal_splits}")
    for i, idx in enumerate(ideal_splits):
        chars_at_split = element_counts[idx] if idx < len(element_counts) else total_count
        log_manager.log(f"  分割点{i+1}: 元素[{idx}]/{len(elements)}, 累计字数: {chars_at_split:,}")

    # 生成带缓冲的分割范围（基于字数精确计算缓冲区，而非全局平均）
    split_ranges = []
    for part_idx in range(num_parts):
        if part_idx == 0:
            start = 0
            if ideal_splits:
                end = _find_buffer_end(element_counts, ideal_splits[0], buffer_chars, 'right')
            else:
                end = len(elements)
        elif part_idx == num_parts - 1:
            start = _find_buffer_end(element_counts, ideal_splits[-1], buffer_chars, 'left')
            end = len(elements)
        else:
            start = _find_buffer_end(element_counts, ideal_splits[part_idx - 1], buffer_chars, 'left')
            end = _find_buffer_end(element_counts, ideal_splits[part_idx], buffer_chars, 'right')

        # 安全裁剪
        start = max(0, min(start, len(elements) - 1))
        end = max(start + 1, min(end, len(elements)))
        split_ranges.append((start, end))

    for i, (s, e) in enumerate(split_ranges):
        part_chars = element_counts[min(e, len(element_counts)) - 1] - (element_counts[s - 1] if s > 0 else 0) if e > s else 0
        log_manager.log(f"  Part{i+1}: 元素[{s}:{e}], 约 {part_chars:,} 字")

    # 生成分割后的文件
    generated_files = []
    part_info = []

    for i, (start_idx, end_idx) in enumerate(split_ranges):
        part_num = i + 1
        dest_filename = f"{base_name}_Part{part_num}.docx"
        dest_path = os.path.join(output_dir, dest_filename)

        shutil.copy2(src_path, dest_path)
        doc_copy = Document(dest_path)

        total_elems = len(get_all_content_elements(doc_copy))
        delete_elements_in_range(doc_copy, end_idx, total_elems + 5000)
        delete_elements_in_range(doc_copy, 0, start_idx)
        doc_copy.save(dest_path)

        first_text = extract_text_from_elements(elements, start_idx, min(start_idx + 3, end_idx))
        last_text = extract_text_from_elements(elements, max(start_idx, end_idx - 3), end_idx)

        part_info.append({
            'path': dest_path,
            'first_anchor': first_text[:200] if first_text else "",
            'last_anchor': last_text[-200:] if last_text else "",
            'start_idx': start_idx,
            'end_idx': end_idx
        })

        generated_files.append(dest_path)
        log_manager.log(f"生成: {dest_filename}")

    return generated_files, part_info, element_ratios


def find_element_index_by_char_count(element_counts, target_chars):
    """二分查找：根据目标字数找到最接近的元素索引

    修复：原逻辑只返回累计字数 < target 的最后一个元素，
    当存在超大元素（如大表格）时，多个目标值会映射到同一索引。
    现在额外比较下一个元素，返回累计字数更接近 target 的那个。
    """
    left, right = 0, len(element_counts) - 1
    best_idx = 0
    while left <= right:
        mid = (left + right) // 2
        if element_counts[mid] < target_chars:
            best_idx = mid
            left = mid + 1
        else:
            right = mid - 1
    # 检查下一个元素是否更接近目标字数
    if best_idx + 1 < len(element_counts):
        diff_before = target_chars - element_counts[best_idx]
        diff_after = element_counts[best_idx + 1] - target_chars
        if diff_after < diff_before:
            return best_idx + 1
    return best_idx


def _find_buffer_end(element_counts, split_idx, buffer_chars, direction='right'):
    """基于字数精确计算缓冲区边界（而非用全局平均估算元素数）

    direction='right': 从 split_idx 向右扩展 buffer_chars 字，返回结束索引
    direction='left':  从 split_idx 向左扩展 buffer_chars 字，返回开始索引
    """
    n = len(element_counts)
    if n == 0:
        return split_idx

    if direction == 'right':
        base_chars = element_counts[split_idx] if split_idx < n else element_counts[-1]
        target = base_chars + buffer_chars
        for i in range(split_idx + 1, n):
            if element_counts[i] >= target:
                return i + 1  # +1 因为 end 是开区间
        return n
    else:  # left
        base_chars = element_counts[split_idx] if split_idx < n else element_counts[-1]
        target = base_chars - buffer_chars
        if target <= 0:
            return 0
        for i in range(split_idx - 1, -1, -1):
            if element_counts[i] <= target:
                return i
        return 0