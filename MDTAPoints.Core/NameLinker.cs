namespace MDTAPoints.Core;

public sealed record KnownStudent(
    string Id,
    string SchoolId,
    string NormalizedName,
    IReadOnlyList<string> Aliases);

public sealed record LinkDecision(LinkStatus Status, string? StudentId, string? SuggestedStudentId);

public static class NameNormalizer
{
    public static string Normalize(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return "";
        }

        var text = name.Trim();
        var comma = text.IndexOf(',');
        if (comma > 0 && comma < text.Length - 1)
        {
            var last = text[..comma];
            var first = text[(comma + 1)..];
            text = first + " " + last;
        }

        var chars = text.ToLowerInvariant().Select(character => char.IsLetterOrDigit(character) ? character : ' ').ToArray();
        return string.Join(' ', new string(chars).Split(' ', StringSplitOptions.RemoveEmptyEntries));
    }

    public static IReadOnlyList<string> SplitCompetitors(string entryName)
    {
        var parts = entryName.Split(" & ", StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        return parts.Length == 0 ? new[] { entryName.Trim() } : parts;
    }
}

public static class NameLinker
{
    public static LinkDecision Classify(string rawName, string schoolId, IEnumerable<KnownStudent> students)
    {
        var normalized = NameNormalizer.Normalize(rawName);
        var atSchool = students.Where(student => student.SchoolId == schoolId).ToList();

        var exact = atSchool.FirstOrDefault(student =>
            student.NormalizedName == normalized || student.Aliases.Contains(normalized));
        if (exact is not null)
        {
            return new LinkDecision(LinkStatus.Linked, exact.Id, exact.Id);
        }

        KnownStudent? best = null;
        var bestDistance = int.MaxValue;
        foreach (var student in atSchool)
        {
            Consider(normalized, student.NormalizedName, student);
            foreach (var alias in student.Aliases)
            {
                Consider(normalized, alias, student);
            }
        }

        if (best is not null)
        {
            return new LinkDecision(LinkStatus.Possible, null, best.Id);
        }

        return new LinkDecision(LinkStatus.New, null, null);

        void Consider(string incoming, string existing, KnownStudent student)
        {
            var distance = Similarity(incoming, existing);
            if (distance < bestDistance)
            {
                bestDistance = distance;
                best = student;
            }
        }
    }

    public static bool CanPublish(IEnumerable<LinkStatus> statuses) =>
        statuses.All(status => status != LinkStatus.Possible);

    private static int Similarity(string incoming, string existing)
    {
        if (incoming.Length == 0 || existing.Length == 0 || incoming == existing)
        {
            return incoming == existing ? 0 : int.MaxValue;
        }

        var incomingParts = incoming.Split(' ');
        var existingParts = existing.Split(' ');
        var incomingLast = incomingParts[^1];
        var existingLast = existingParts[^1];
        if (incomingLast == existingLast && FirstNamesSimilar(incomingParts[0], existingParts[0]))
        {
            return Levenshtein(incomingParts[0], existingParts[0]);
        }

        var full = Levenshtein(incoming, existing);
        return full <= 2 ? full : int.MaxValue;
    }

    private static bool FirstNamesSimilar(string left, string right)
    {
        if (left == right)
        {
            return true;
        }

        var shorter = left.Length <= right.Length ? left : right;
        var longer = left.Length <= right.Length ? right : left;
        if (shorter.Length >= 3 && longer.StartsWith(shorter, StringComparison.Ordinal))
        {
            return true;
        }

        return Levenshtein(left, right) <= 2;
    }

    public static int Levenshtein(string left, string right)
    {
        var previous = new int[right.Length + 1];
        var current = new int[right.Length + 1];
        for (var column = 0; column <= right.Length; column++)
        {
            previous[column] = column;
        }

        for (var row = 1; row <= left.Length; row++)
        {
            current[0] = row;
            for (var column = 1; column <= right.Length; column++)
            {
                var cost = left[row - 1] == right[column - 1] ? 0 : 1;
                current[column] = Math.Min(
                    Math.Min(current[column - 1] + 1, previous[column] + 1),
                    previous[column - 1] + cost);
            }

            (previous, current) = (current, previous);
        }

        return previous[right.Length];
    }
}
