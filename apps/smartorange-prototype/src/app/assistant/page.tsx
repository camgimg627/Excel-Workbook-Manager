"use client";

import * as React from "react";
import { PageHeader, Surface } from "@/components/ui";
import { useDemoSession } from "@/lib/session-context";

const promptSuggestions = [
  "What is running low this week?",
  "Do we have enough Olive Oil for dinner?",
  "When did we last buy Paper Towels?",
  "What expires soon?",
  "Where is Shampoo right now?",
];

export default function AssistantPage() {
  const { activeHousehold, activeMember, state, actions } = useDemoSession();
  const [prompt, setPrompt] = React.useState("");
  const messages = state.assistantMessages.filter((entry) => entry.householdId === activeHousehold.id);

  return (
    <div className="page">
      <PageHeader
        eyebrow="Household assistant"
        title="Grounded answers over the live demo state."
        description={`The assistant stays household-aware. Right now it is speaking as the ${activeMember.role} view for ${activeHousehold.name}.`}
      />

      <div className="twoColumn">
        <Surface strong>
          <div className="sectionHeading">
            <h3 className="sectionTitle">Suggested prompts</h3>
            <span className="badge badgeGreen">No external AI calls</span>
          </div>
          <div className="cluster">
            {promptSuggestions.map((entry) => (
              <button key={entry} type="button" className="buttonGhost" onClick={() => setPrompt(entry)}>
                {entry}
              </button>
            ))}
          </div>
          <form
            className="stack"
            onSubmit={(event) => {
              event.preventDefault();
              if (!prompt.trim()) {
                return;
              }
              actions.askAssistant(prompt.trim());
              setPrompt("");
            }}
          >
            <label className="field">
              <span className="label">Ask about inventory, shopping, expiry, or location</span>
              <textarea className="textarea" value={prompt} onChange={(event) => setPrompt(event.target.value)} />
            </label>
            <button className="button" type="submit">
              Send prompt
            </button>
          </form>
        </Surface>

        <Surface>
          <div className="sectionHeading">
            <h3 className="sectionTitle">Conversation</h3>
            <span className="badge badgeNeutral">{messages.length} messages</span>
          </div>
          <div className="chatColumn">
            {messages.map((message) => (
              <div
                key={message.id}
                className={`chatBubble ${message.role === "user" ? "chatBubbleUser" : "chatBubbleAssistant"}`}
              >
                {message.content}
              </div>
            ))}
          </div>
        </Surface>
      </div>
    </div>
  );
}
