package com.github.abcinc.docx;

import static com.google.common.collect.Iterables.getFirst;
import static com.google.common.collect.Iterables.getLast;
import static com.google.common.collect.Iterables.isEmpty;
import static java.util.Objects.requireNonNull;
import static java.util.stream.Collectors.joining;
import static java.util.stream.Collectors.toUnmodifiableSet;

import com.google.common.base.Splitter;
import com.google.common.collect.Range;
import com.google.common.primitives.Ints;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Stream;

public class Slicing {

  public static List<Range<Integer>> calcPages(List<String> pages) {
    if (isEmpty(pages)) {
      return List.of();
    }

    Splitter splitter = Splitter.on(',').omitEmptyStrings().trimResults();
    Splitter rangeSplitter = Splitter.on('-').trimResults();

    List<String> items = splitter.splitToList(pages.stream().collect(joining(",")));
    List<Range<Integer>> ranges = new ArrayList<>();
    for (String item : items) {
      List<String> fromTo = List.of(item, item);
      if (item.contains("-")) {
        fromTo = rangeSplitter.splitToList(item);
      }
      Integer lower = requireNonNull(Ints.tryParse(getFirst(fromTo, null)));
      Integer upper = requireNonNull(Ints.tryParse(getLast(fromTo)));
      ranges.add(Range.closed(lower, upper));
    }
    if (ranges.size() < 2) {
      return ranges;
    }

    ranges.sort(Comparator.comparing(Range::lowerEndpoint));

    List<Range<Integer>> mergedRanges = new ArrayList<>();
    Range<Integer> currentRange = ranges.get(0);
    for (int i = 1; i < ranges.size(); i++) {
      Range<Integer> other = ranges.get(i);
      if (currentRange.upperEndpoint() >= other.lowerEndpoint() - 1) {
        currentRange = currentRange.span(other);
      } else {
        mergedRanges.add(currentRange);
        currentRange = other;
      }
    }
    mergedRanges.add(currentRange);
    return mergedRanges;
  }

  public static Collection<Integer> rangesToList(List<Range<Integer>> ranges) {
    return ranges.stream()
        .flatMap(r -> Stream.iterate(r.lowerEndpoint(), r, n -> n + 1))
        .collect(toUnmodifiableSet());
  }
}
