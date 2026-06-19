/**
 * Removing a reply by ID must also remove it from the parent's replies
 * array.
 *
 * CommentManager tracks replies both as first-class map entries and inside
 * the parent entry's replies array. removeComment(replyId) used to delete
 * only the map entry, so getReplies()/hasReplies()/getCommentThread() kept
 * returning the deleted reply while generateCommentsXml() (map-driven)
 * correctly excluded it — two disagreeing views of the same thread.
 */
import { Document } from '../../src/core/Document';

describe('C87: removeComment(replyId) keeps parent replies in sync', () => {
  it('clears the reply from the parent thread views', () => {
    const doc = Document.create();
    try {
      const parent = doc.createComment('User1', 'Parent comment');
      const reply = doc.createReply(parent.getId(), 'User2', 'Reply comment');
      const manager = doc.getCommentManager();

      expect(manager.getReplies(parent.getId())).toHaveLength(1);
      expect(manager.hasReplies(parent.getId())).toBe(true);

      expect(doc.removeComment(reply.getId())).toBe(true);

      expect(manager.getComment(reply.getId())).toBeUndefined();
      expect(manager.getReplies(parent.getId())).toEqual([]);
      expect(manager.hasReplies(parent.getId())).toBe(false);
      expect(manager.getCommentThread(parent.getId())?.replies).toEqual([]);

      // The parent itself is untouched
      expect(manager.getComment(parent.getId())).toBe(parent);
    } finally {
      doc.dispose();
    }
  });

  it('leaves sibling replies intact when one reply is removed', () => {
    const doc = Document.create();
    try {
      const parent = doc.createComment('User1', 'Parent comment');
      const reply1 = doc.createReply(parent.getId(), 'User2', 'First reply');
      const reply2 = doc.createReply(parent.getId(), 'User3', 'Second reply');
      const manager = doc.getCommentManager();

      expect(doc.removeComment(reply1.getId())).toBe(true);

      const remaining = manager.getReplies(parent.getId());
      expect(remaining).toHaveLength(1);
      expect(remaining[0]).toBe(reply2);
      expect(manager.getCommentThread(parent.getId())?.replies).toEqual([reply2]);
    } finally {
      doc.dispose();
    }
  });

  it('still removes a whole thread when the parent is removed', () => {
    const doc = Document.create();
    try {
      const parent = doc.createComment('User1', 'Parent comment');
      const reply = doc.createReply(parent.getId(), 'User2', 'Reply comment');
      const manager = doc.getCommentManager();

      expect(doc.removeComment(parent.getId())).toBe(true);

      expect(manager.getComment(parent.getId())).toBeUndefined();
      expect(manager.getComment(reply.getId())).toBeUndefined();
      expect(manager.getCount()).toBe(0);
    } finally {
      doc.dispose();
    }
  });
});
