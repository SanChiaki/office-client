import { SessionSidebar } from "../../src/components/SessionSidebar";

export function sessionSidebarEmptyStateContract() {
  return (
    <SessionSidebar
      sessions={[]}
      onCreateSession={() => {}}
      onSelectSession={() => {}}
      onDeleteSession={() => {}}
    />
  );
}
