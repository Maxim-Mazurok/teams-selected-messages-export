import type { MessageSnapshot } from "./types.js";
import { toComparableTime, toComparableNumericId } from "./utilities.js";

export function getMessageRange(
  messageIds: string[],
  anchorId: string,
  targetId: string
): string[] {
  const anchorIndex = messageIds.indexOf(anchorId);
  const targetIndex = messageIds.indexOf(targetId);

  if (anchorIndex === -1 || targetIndex === -1) {
    return [targetId];
  }

  const start = Math.min(anchorIndex, targetIndex);
  const end = Math.max(anchorIndex, targetIndex);
  return messageIds.slice(start, end + 1);
}

export function orderMessagesForExport(messages: MessageSnapshot[]): MessageSnapshot[] {
  return [...messages].sort((left, right) => {
    const leftTime = toComparableTime(left.dateTime);
    const rightTime = toComparableTime(right.dateTime);
    if (leftTime !== null && rightTime !== null && leftTime !== rightTime) {
      return leftTime - rightTime;
    }

    const leftId = toComparableNumericId(left.id);
    const rightId = toComparableNumericId(right.id);
    if (leftId !== null && rightId !== null && leftId !== rightId) {
      return leftId - rightId;
    }

    const leftFallback = left.captureOrder ?? left.index ?? 0;
    const rightFallback = right.captureOrder ?? right.index ?? 0;
    return leftFallback - rightFallback;
  });
}
