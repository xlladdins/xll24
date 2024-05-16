// excel_clock.h - Excel clock
#pragma once
#include <chrono>

namespace xll {

	constexpr auto excel_epoch = std::chrono::sys_days{ std::chrono::year{1900} / std::chrono::January / 0 };
	
	constexpr auto to_days(double d)
	{
		return std::chrono::sys_days{ excel_epoch + std::chrono::days{static_cast<int>(d)} };
	}
	constexpr auto to_excel(const std::chrono::sys_days& sys)
	{
		return std::chrono::duration_cast<std::chrono::days>(sys - excel_epoch).count();
	}
#if 0
	// https://stackoverflow.com/questions/33964461/handling-julian-days-in-c11-14
	struct excel_clock;
    template <class Dur>
    using excel_time = std::chrono::time_point<excel_clock, Dur>;

	// Excel clock represented as days since 1900-01-01.
	struct excel_clock
	{
		using rep = double;
		using period = std::chrono::days::period;
		using duration = std::chrono::duration<rep, period>;
		using time_point = std::chrono::time_point<excel_clock>;

		static constexpr bool is_steady = false;

		static time_point now() noexcept
		{
			using namespace std::chrono;

			return clock_cast<excel_clock>(system_clock::now());
		}

		template <class Dur>
		static auto from_sys(std::chrono::sys_time<Dur> const& tp) noexcept
		{
			using namespace std::chrono;
			static auto epoch = sys_days{ January / 0 / 1900 };

			return excel_time{ tp - epoch };
		}

		template <class Dur>
		static auto to_sys(excel_time<Dur> const& tp) noexcept
		{
			using namespace std::chrono;
			
			return sys_time{ tp - clock_cast<excel_clock>(sys_days{}) };
		}
	};

	inline auto to_time(double d)
	{
            return /*std::chrono::floor<std::chrono::days>*/(excel_time { excel_clock::duration { d } });
	}

	inline std::chrono::sys_days to_days(double d)
	{
		using namespace std::chrono;

		return time_point_cast<days>(excel_clock::to_sys(to_time(d)));
	}

	inline std::chrono::year_month_day to_ymd(double d)
	{
		return std::chrono::year_month_day{ to_days(d) };
	}

	inline double from_ymd(const std::chrono::year_month_day& d)
	{
		return excel_clock::from_sys(std::chrono::sys_days{ d }).time_since_epoch().count();
	}
#endif // 0
}