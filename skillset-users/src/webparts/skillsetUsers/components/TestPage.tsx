import * as React from 'react';
import { useState, useEffect } from 'react';
import { sp } from "@pnp/sp/presets/all";
import { DefaultButton, PrimaryButton, Stack, Text, MessageBar, MessageBarType } from '@fluentui/react';
import { IDropdownOption } from '@fluentui/react';

interface ITestPageProps {
  selectedSkillIds: number[];
  skillsetOptions: IDropdownOption[];
  welcomeName: string;
  userEmail: string;
  onLogout: () => void;
  onBack: () => void;
}

interface IQuestion {
  Id: number;
  Title: string;
  OptionA: string;
  OptionB: string;
  OptionC: string;
  OptionD: string;
  CorrectAnswer: string;
}

interface ISavedTest {
  Id: number;
  SkillId: number;
  SkillName: string;
  QuestionsJson: string;
  AnswersJson: string;
}

const TestPage: React.FC<ITestPageProps> = ({ selectedSkillIds, skillsetOptions, welcomeName, userEmail, onLogout, onBack }) => {
  const [saveDraftMessage, setSaveDraftMessage] = useState<string>('');
  const [questions, setQuestions] = useState<IQuestion[]>([]);
  const [selectedSkillId, setSelectedSkillId] = useState<number | null>(null);
  const [answers, setAnswers] = useState<{ [key: number]: string }>({});
  const [score, setScore] = useState<number | null>(null);
  const [submitted, setSubmitted] = useState(false);
  const [previousResults, setPreviousResults] = useState<any[]>([]);
  const [showResults, setShowResults] = useState(false);
  const [savedDrafts, setSavedDrafts] = useState<ISavedTest[]>([]);
  const [currentDraftId, setCurrentDraftId] = useState<number | null>(null);

  const availableSkills = skillsetOptions.filter(opt => selectedSkillIds.includes(Number(opt.key)));

  useEffect(() => {
    fetchSavedDrafts();
  }, []);

  const fetchSavedDrafts = async () => {
    try {
      const items = await sp.web.lists.getByTitle("Saved_Tests")
        .items
        .filter(`Title eq '${welcomeName}'`)
        .select("Id", "SkillId", "SkillName", "QuestionsJson", "AnswersJson")
        .get();
      setSavedDrafts(items);
    } catch (error) {
      console.error("Error fetching saved drafts:", error);
    }
  };

  const fetchQuestions = async (skillId: number) => {
    try {
      const items: IQuestion[] = await sp.web.lists.getByTitle("Skillset_Questions")
        .items
        .filter(`SkillsetId eq ${skillId}`)
        .top(100)
        .get();
      const shuffled = items.sort(() => 0.5 - Math.random()).slice(0, 10);
      setQuestions(shuffled);
      setSelectedSkillId(skillId);
      setAnswers({});
      setScore(null);
      setSubmitted(false);
      setShowResults(false);
      setCurrentDraftId(null);
    } catch (error) {
      console.error("Error fetching questions:", error);
    }
  };

  const fetchPreviousResults = async () => {
    try {
      const items = await sp.web.lists.getByTitle("Test_Results")
        .items
        .filter(`Email eq '${userEmail}'`)
        .select("SkillName", "Score", "Passed_x003f_", "DateTaken")
        .orderBy("DateTaken", false)
        .get();
      setPreviousResults(items);
      setShowResults(true);
      setSelectedSkillId(null);
    } catch (error) {
      console.error("Error fetching previous results:", error);
    }
  };

  const handleAnswerChange = (questionId: number, selectedOptionKey: string) => {
    setAnswers(prev => ({ ...prev, [questionId]: selectedOptionKey }));
  };

  const handleSaveDraft = async () => {
    try {
      const payload = {
        Title: welcomeName,
        SkillId: selectedSkillId,
        SkillName: skillsetOptions.find(opt => opt.key === selectedSkillId)?.text || '',
        QuestionsJson: JSON.stringify(questions),
        AnswersJson: JSON.stringify(answers),
        DateSaved: new Date().toISOString()
      };

      if (currentDraftId !== null) {
        await sp.web.lists.getByTitle("Saved_Tests").items.getById(currentDraftId).update(payload);
      } else {
        await sp.web.lists.getByTitle("Saved_Tests").items.add(payload);
      }

      setQuestions([]);
      setSelectedSkillId(null);
      setAnswers({});
      setCurrentDraftId(null);
      fetchSavedDrafts();
      setSaveDraftMessage("💾 This skill test is now saved as a draft. You can continue it later.");
    } catch (error) {
      console.error("Error saving draft:", error);
    }
  };

  const handleLoadDraft = (draft: ISavedTest) => {
    setQuestions(JSON.parse(draft.QuestionsJson));
    setAnswers(JSON.parse(draft.AnswersJson));
    setSelectedSkillId(draft.SkillId);
    setScore(null);
    setSubmitted(false);
    setShowResults(false);
    setCurrentDraftId(draft.Id);
  };

  const handleSubmit = async () => {
    let correct = 0;
    questions.forEach(q => {
      if (answers[q.Id] === q.CorrectAnswer) correct++;
    });

    const passed = correct >= 6 ? "Yes" : "No";
    setScore(correct);
    setSubmitted(true);

    try {
      await sp.web.lists.getByTitle("Test_Results").items.add({
        Title: welcomeName,
        Email: userEmail,
        SkillName: skillsetOptions.find(opt => opt.key === selectedSkillId)?.text || '',
        Score: correct,
        "Passed_x003f_": passed,
        DateTaken: new Date().toISOString()
      });

      if (currentDraftId !== null) {
        await sp.web.lists.getByTitle("Saved_Tests").items.getById(currentDraftId).recycle();
        fetchSavedDrafts();
        setCurrentDraftId(null);
      }
    } catch (error) {
      console.error("Error saving test result:", error);
    }
  };

  return (
    <>
      <div style={{ padding: 20 }}>
        <Stack horizontal horizontalAlign="space-between">
          <Text variant="xxLarge">🧪 Skillset Dashboard</Text>
          <PrimaryButton
            text={showResults ? "⬅ Back to Skill Test Dashboard" : "View Completed Tests"}
            onClick={() => !showResults ? fetchPreviousResults() : setShowResults(false)}
          />
        </Stack>

        {showResults && (
          <div style={{ marginTop: 20 }}>
            <Text variant="xLarge">📚 Completed Tests</Text>
            {previousResults.length === 0 ? (
              <Text>No previous tests found.</Text>
            ) : (
              <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: 20, fontSize: 14 }}>
                <thead>
                  <tr>
                    <th style={{ textAlign: 'left', padding: '8px 12px' }}>Skill</th>
                    <th style={{ textAlign: 'center', padding: '8px 12px' }}>Score</th>
                    <th style={{ textAlign: 'center', padding: '8px 12px' }}>Result</th>
                    <th style={{ textAlign: 'right', padding: '8px 12px' }}>Date</th>
                  </tr>
                </thead>
                <tbody>
                  {previousResults.map((item, index) => (
                    <tr key={index}>
                      <td style={{ padding: '8px 12px' }}>{item.SkillName}</td>
                      <td style={{ textAlign: 'center', padding: '8px 12px' }}>{item.Score}/10</td>
                      <td style={{ textAlign: 'center', padding: '8px 12px' }}>
                        {item["Passed_x003f_"] === "Yes" ? "Passed" : "Failed"}
                      </td>
                      <td style={{ textAlign: 'right', padding: '8px 12px' }}>
                        {new Date(item.DateTaken).toLocaleString()}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        )}

        {savedDrafts.length > 0 && !selectedSkillId && !showResults && (
          <div style={{ marginTop: 30 }}>
            <Text variant="xLarge">📌 Saved Test Drafts ({savedDrafts.length})</Text>
            <Stack horizontal wrap tokens={{ childrenGap: 10 }} style={{ marginTop: 10 }}>
              {savedDrafts.map(d => (
                <PrimaryButton key={d.Id} text={d.SkillName} onClick={() => handleLoadDraft(d)} />
              ))}
            </Stack>
          </div>
        )}

        {saveDraftMessage && (
          <MessageBar
            messageBarType={MessageBarType.success}
            isMultiline={false}
            styles={{ root: { marginTop: 30 }, icon: { display: 'none' } }}
          >
            <span style={{ fontSize: 15 }}>{saveDraftMessage}</span>
          </MessageBar>
        )}

        {!selectedSkillId && !showResults && (
          <div style={{ marginTop: 30 }}>
            <Text variant="xLarge">🧠 Your Selected Skills</Text>
            <Stack horizontal wrap tokens={{ childrenGap: 10 }} style={{ marginTop: 30 }}>
              {availableSkills.map(skill => (
                <PrimaryButton
                  key={skill.key}
                  text={skill.text}
                  onClick={() => fetchQuestions(Number(skill.key))}
                />
              ))}
            </Stack>
          </div>
        )}

        {selectedSkillId && questions.length > 0 && (
          <div style={{ marginTop: 30 }}>
            <Text variant="xLarge">
              Skill Test: {skillsetOptions.find(opt => opt.key === selectedSkillId)?.text}
            </Text>

            {questions.map((q, index) => (
              <div key={q.Id} style={{ marginBottom: 25, padding: 10, border: '1px solid #ccc', borderRadius: 8 }}>
                <Text variant="large">Q{index + 1}. {q.Title}</Text>
                <Stack tokens={{ childrenGap: 8 }} style={{ marginTop: 10 }}>
                  {(['OptionA', 'OptionB', 'OptionC', 'OptionD'] as const).map(optionKey => {
                    const keyLetter = optionKey.replace('Option', '');
                    const optionValue = q[optionKey];
                    const isSelected = answers[q.Id] === keyLetter;
                    const isCorrect = q.CorrectAnswer === keyLetter;

                    let backgroundColor = '';
                    if (submitted) {
                      if (isCorrect) backgroundColor = '#d4edda';
                      else if (isSelected && !isCorrect) backgroundColor = '#f8d7da';
                    }

                    return (
                      <label key={optionKey} style={{
                        display: 'block',
                        cursor: submitted ? 'default' : 'pointer',
                        marginBottom: 6,
                        backgroundColor,
                        padding: 6,
                        borderRadius: 4
                      }}>
                        <input
                          type="radio"
                          name={`question-${q.Id}`}
                          value={keyLetter}
                          checked={isSelected}
                          disabled={submitted}
                          onChange={() => handleAnswerChange(q.Id, keyLetter)}
                          style={{ marginRight: 8 }}
                        />
                        {optionValue}
                        {submitted && isSelected && (
                          <strong style={{ marginLeft: 10 }}>
                            {isCorrect ? '✅ Correct' : '❌ Your Answer'}
                          </strong>
                        )}
                        {submitted && !isSelected && isCorrect && (
                          <strong style={{ marginLeft: 10, color: 'green' }}>
                            ✔ Correct Answer
                          </strong>
                        )}
                      </label>
                    );
                  })}
                </Stack>
              </div>
            ))}

            {!submitted && (
              <Stack horizontal tokens={{ childrenGap: 10 }}>
                <PrimaryButton text="Submit" onClick={handleSubmit} />
                <DefaultButton text="Save for Later" onClick={handleSaveDraft} />
                <DefaultButton text="Back to Skills" onClick={() => {
                  setSelectedSkillId(null);
                  setQuestions([]);
                  setCurrentDraftId(null);
                }} />
              </Stack>
            )}

            {submitted && score !== null && (
              <MessageBar
                messageBarType={score >= 6 ? MessageBarType.success : MessageBarType.error}
                isMultiline={false}
                styles={{
                  icon: { display: 'none' },
                  root: {
                    marginTop: 20,
                    fontWeight: 500,
                    alignItems: 'center',
                    paddingLeft: 12
                  },
                  content: { display: 'flex', alignItems: 'center' }
                }}
              >
                <span style={{ fontSize: '15px', display: 'flex', alignItems: 'center' }}>
                  <span role="img" aria-label="score" style={{ marginRight: 6 }}>🎯</span>
                  Your score: <strong>{score}/10</strong>. You have
                  <strong style={{ margin: '0 4px' }}>{score >= 6 ? 'passed ✅' : 'failed ❌'}</strong>
                  the test.
                </span>
              </MessageBar>
            )}
          </div>
        )}

        <DefaultButton text="⬅ Back to Dashboard" onClick={onBack} style={{ marginTop: 30 }} />
      </div>
    </>
  );
};

export default TestPage;
