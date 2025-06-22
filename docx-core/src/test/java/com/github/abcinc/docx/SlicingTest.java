package com.github.abcinc.docx;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrowsExactly;

import com.google.common.collect.Range;
import java.util.List;
import org.junit.jupiter.api.Test;

public class SlicingTest {
  @Test
  public void testRanges() {
    List<Range<Integer>> ranges = Slicing.calcPages(List.of("1", "4,5", "7-7,9-11", "4-5,1-2"));
    assertEquals(4, ranges.size());
    assertEquals(Range.closed(1, 2), ranges.get(0));
    assertEquals(Range.closed(4, 5), ranges.get(1));
    assertEquals(Range.closed(7, 7), ranges.get(2));
    assertEquals(Range.closed(9, 11), ranges.get(3));
  }

  @Test
  public void testInvalidRange() {
    assertThrowsExactly(IllegalArgumentException.class, () -> Slicing.calcPages(List.of("3-2")));
  }
}
