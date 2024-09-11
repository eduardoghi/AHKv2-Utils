IsWorkDay(date) {
    weekday := FormatTime(date, "WDay")

    if (weekday = 1 || weekday = 7) {
        return false
    }

    return true
}
