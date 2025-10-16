/**
 * Examples of using Comments in DocXML
 *
 * This file demonstrates:
 * - Simple comments on paragraphs
 * - Multiple comments with different authors
 * - Comment replies (threaded comments)
 * - Comments with formatted content
 * - Document review workflow
 */

import { Document, Run, Comment } from '../../src/index';

/**
 * Example 1: Simple Comments
 * Add basic comments to paragraphs
 */
async function example1_simpleComments() {
  console.log('\nExample 1: Simple Comments');

  const doc = Document.create();

  // Add document title
  doc.createParagraph('Document Review')
    .setStyle('Title')
    .setAlignment('center');

  // Create first paragraph with comment
  const para1 = doc.createParagraph(
    'This is the first paragraph that needs to be reviewed.'
  );
  const comment1 = doc.createComment(
    'John Smith',
    'This paragraph looks good!',
    'JS'
  );
  para1.addComment(comment1);

  // Create second paragraph with comment
  const para2 = doc.createParagraph(
    'This paragraph contains important information about our project timeline.'
  );
  const comment2 = doc.createComment(
    'Jane Doe',
    'Can we update the timeline to reflect recent changes?',
    'JD'
  );
  para2.addComment(comment2);

  // Create third paragraph without comment
  doc.createParagraph('This paragraph has no comments.');

  // Save document
  await doc.save('examples/11-comments/example1-simple-comments.docx');
  console.log('Created: example1-simple-comments.docx');
  console.log(`   Comments: ${doc.getCommentStats().total}`);
}

/**
 * Example 2: Multiple Authors
 * Demonstrate comments from different team members
 */
async function example2_multipleAuthors() {
  console.log('\nExample 2: Multiple Authors');

  const doc = Document.create({
    properties: {
      title: 'Team Review Document',
      creator: 'Project Team'
    }
  });

  // Add heading
  doc.createParagraph('Quarterly Report - Draft')
    .setStyle('Heading1')
    .setAlignment('center');

  // Section 1
  doc.createParagraph('Executive Summary')
    .setStyle('Heading2');

  const execSummary = doc.createParagraph(
    'Our Q4 results show significant growth across all departments.'
  );
  const comment1 = doc.createComment(
    'Alice Johnson',
    'Great work! Can we add specific percentage numbers?',
    'AJ'
  );
  execSummary.addComment(comment1);

  // Section 2
  doc.createParagraph('Financial Performance')
    .setStyle('Heading2');

  const financial = doc.createParagraph(
    'Revenue increased substantially compared to last quarter.'
  );
  const comment2 = doc.createComment(
    'Bob Williams',
    'Please include the exact revenue figures here.',
    'BW'
  );
  financial.addComment(comment2);

  // Section 3
  doc.createParagraph('Future Outlook')
    .setStyle('Heading2');

  const outlook = doc.createParagraph(
    'We expect continued growth in the coming year.'
  );
  const comment3 = doc.createComment(
    'Carol Martinez',
    'This is too vague. We need concrete projections.',
    'CM'
  );
  outlook.addComment(comment3);

  const comment4 = doc.createComment(
    'David Chen',
    'I agree with Carol. Let\'s schedule a meeting to discuss Q1 targets.',
    'DC'
  );
  outlook.addComment(comment4);

  // Save document
  await doc.save('examples/11-comments/example2-multiple-authors.docx');
  console.log('Created: example2-multiple-authors.docx');
  const stats = doc.getCommentStats();
  console.log(`   Comments: ${stats.total} from ${stats.authors.length} authors`);
  console.log(`   Authors: ${stats.authors.join(', ')}`);
}

/**
 * Example 3: Comment Replies
 * Demonstrate threaded comments with replies
 */
async function example3_commentReplies() {
  console.log('\nExample 3: Comment Replies (Threaded Comments)');

  const doc = Document.create();

  // Add title
  doc.createParagraph('Project Proposal')
    .setStyle('Title');

  // Add content paragraph
  const proposal = doc.createParagraph(
    'We propose implementing a new customer feedback system using advanced sentiment analysis.'
  );

  // Create initial comment
  const mainComment = doc.createComment(
    'Sarah Thompson',
    'This is an interesting idea. What\'s the estimated budget?',
    'ST'
  );
  proposal.addComment(mainComment);

  // Create reply to the comment
  doc.createReply(
    mainComment.getId(),
    'Mike Rodriguez',
    'Good question! I estimate around $50,000 for the initial implementation.',
    'MR'
  );

  // Create another reply
  doc.createReply(
    mainComment.getId(),
    'Sarah Thompson',
    'Thanks Mike! That seems reasonable. Let\'s discuss this in the next meeting.',
    'ST'
  );

  // Add another paragraph with threaded discussion
  const timeline = doc.createParagraph(
    'The project timeline is estimated at 6 months from kickoff to deployment.'
  );

  const timelineComment = doc.createComment(
    'Lisa Brown',
    'Can we shorten this to 4 months? We have a tight deadline.',
    'LB'
  );
  timeline.addComment(timelineComment);

  doc.createReply(
    timelineComment.getId(),
    'Mike Rodriguez',
    'That would require additional resources. Let me check with the team.',
    'MR'
  );

  doc.createReply(
    timelineComment.getId(),
    'Lisa Brown',
    'Please do. We might be able to allocate more developers if needed.',
    'LB'
  );

  // Save document
  await doc.save('examples/11-comments/example3-comment-replies.docx');
  console.log('Created: example3-comment-replies.docx');
  const stats = doc.getCommentStats();
  console.log(`   Total comments: ${stats.total}`);
  console.log(`   Top-level: ${stats.topLevel}, Replies: ${stats.replies}`);
}

/**
 * Example 4: Formatted Comments
 * Demonstrate comments with formatted content
 */
async function example4_formattedComments() {
  console.log('\nExample 4: Formatted Comments');

  const doc = Document.create();

  // Add title
  doc.createParagraph('Technical Documentation')
    .setStyle('Heading1');

  // Add code-like paragraph
  const codePara = doc.createParagraph('function calculateTotal(items) { ... }')
    .setStyle('Normal')
    .setLeftIndent(720); // Indent to look like code

  // Create comment with formatted content
  const formattedRuns = [
    new Run('This function needs ').setBold(false),
    new Run('error handling').setBold(true).setColor('FF0000'),
    new Run(' for null values.').setBold(false)
  ];

  const codeComment = Comment.createFormatted(
    'Tech Reviewer',
    formattedRuns,
    'TR'
  );

  // Register and add the comment
  doc.getCommentManager().register(codeComment);
  codePara.addComment(codeComment);

  // Add another paragraph with formatted comment
  const apiPara = doc.createParagraph(
    'The API endpoint returns a JSON response with user data.'
  );

  const apiCommentRuns = [
    new Run('Question: ').setBold(true),
    new Run('What happens if the user is ').setBold(false),
    new Run('not found').setItalic(true),
    new Run('? Should we return 404 or 200 with null?').setBold(false)
  ];

  const apiComment = Comment.createFormatted(
    'API Designer',
    apiCommentRuns,
    'AD'
  );

  doc.getCommentManager().register(apiComment);
  apiPara.addComment(apiComment);

  // Save document
  await doc.save('examples/11-comments/example4-formatted-comments.docx');
  console.log('Created: example4-formatted-comments.docx');
  console.log(`   Comments with formatted content: ${doc.getCommentStats().total}`);
}

/**
 * Example 5: Document Review Workflow
 * Complete document review scenario
 */
async function example5_reviewWorkflow() {
  console.log('\nExample 5: Document Review Workflow');

  const doc = Document.create({
    properties: {
      title: 'Marketing Proposal - Review Draft',
      creator: 'Marketing Team',
      subject: 'Q1 2024 Campaign'
    }
  });

  // Document title
  doc.createParagraph('Q1 2024 Marketing Campaign Proposal')
    .setStyle('Title')
    .setAlignment('center');

  // Section 1: Campaign Overview
  doc.createParagraph('Campaign Overview')
    .setStyle('Heading1');

  const overview = doc.createParagraph(
    'Our Q1 campaign will focus on digital channels with emphasis on social media and influencer partnerships.'
  );
  doc.addCommentToParagraph(
    overview,
    'Marketing Director',
    'Approved. Make sure we track ROI for each channel.',
    'MD'
  );

  // Section 2: Budget
  doc.createParagraph('Budget Allocation')
    .setStyle('Heading1');

  const budget = doc.createParagraph(
    'Total budget: $500,000 divided across social media (40%), influencers (30%), and content creation (30%).'
  );

  const budgetComment = doc.createComment(
    'Finance Manager',
    'The total seems high. Can we reduce to $400,000?',
    'FM'
  );
  budget.addComment(budgetComment);

  doc.createReply(
    budgetComment.getId(),
    'Marketing Director',
    'We can reduce content creation to 20% and save $50,000. New total: $450,000.',
    'MD'
  );

  doc.createReply(
    budgetComment.getId(),
    'Finance Manager',
    'That works. Approved at $450,000.',
    'FM'
  );

  // Section 3: Timeline
  doc.createParagraph('Timeline')
    .setStyle('Heading1');

  const timeline = doc.createParagraph(
    'Campaign launches January 15th and runs through March 31st.'
  );
  doc.addCommentToParagraph(
    timeline,
    'Project Manager',
    '✓ Timeline confirmed. All teams have been notified.',
    'PM'
  );

  // Section 4: Success Metrics
  doc.createParagraph('Success Metrics')
    .setStyle('Heading1');

  const metrics = doc.createParagraph(
    'We will measure success by tracking engagement rates, conversion rates, and brand awareness surveys.'
  );

  const metricsComment = doc.createComment(
    'Data Analyst',
    'Please add specific numerical targets for each metric.',
    'DA'
  );
  metrics.addComment(metricsComment);

  doc.createReply(
    metricsComment.getId(),
    'Marketing Director',
    'Good point. Targets: 5% engagement, 2% conversion, 80% brand awareness.',
    'MD'
  );

  // Final approval section
  doc.createParagraph('\nApproval Status')
    .setStyle('Heading1');

  const approval = doc.createParagraph(
    'This proposal is pending final approval from the executive team.'
  );
  doc.addCommentToParagraph(
    approval,
    'CEO',
    '✓ APPROVED - Great work team! Proceed with the revised $450K budget.',
    'CEO'
  );

  // Save document
  await doc.save('examples/11-comments/example5-review-workflow.docx');
  console.log('Created: example5-review-workflow.docx');
  const stats = doc.getCommentStats();
  console.log(`   Total comments: ${stats.total} (${stats.topLevel} threads, ${stats.replies} replies)`);
  console.log(`   Participants: ${stats.authors.join(', ')}`);
}

/**
 * Run all examples
 */
async function main() {
  console.log('DocXML - Comment Examples\n');
  console.log('═══════════════════════════════════════════════════════════');

  try {
    await example1_simpleComments();
    await example2_multipleAuthors();
    await example3_commentReplies();
    await example4_formattedComments();
    await example5_reviewWorkflow();

    console.log('\n═══════════════════════════════════════════════════════════');
    console.log('All comment examples completed successfully!');
    console.log('\nGenerated files:');
    console.log('   - example1-simple-comments.docx');
    console.log('   - example2-multiple-authors.docx');
    console.log('   - example3-comment-replies.docx');
    console.log('   - example4-formatted-comments.docx');
    console.log('   - example5-review-workflow.docx');
    console.log('\nOpen these files in Microsoft Word to see the comments!');
  } catch (error) {
    console.error('[ERROR]:', error);
    process.exit(1);
  }
}

// Run if executed directly
if (require.main === module) {
  main();
}

export {
  example1_simpleComments,
  example2_multipleAuthors,
  example3_commentReplies,
  example4_formattedComments,
  example5_reviewWorkflow,
};
