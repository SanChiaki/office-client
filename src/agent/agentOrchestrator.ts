export type RouteDecision =
  | {
      mode: "skill";
      skillName: string;
    }
  | {
      mode: "chat";
    };

export function decideRoute(input: string): RouteDecision {
  const trimmed = input.trim();

  if (trimmed.startsWith("/upload_data")) {
    return {
      mode: "skill",
      skillName: "upload_data",
    };
  }

  return {
    mode: "chat",
  };
}
