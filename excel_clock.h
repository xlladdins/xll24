// excel_clock.h - Excel clock
#pragma once
#include <chrono>

namespace xll {

	using std::chrono::operator""y;
	using ymd = std::chrono::year_month_day;

	constexpr const auto epoch = std::chrono::sys_days { std::chrono::December / 30 / 1899 };
	constexpr double to_excel(const std::chrono::sys_days& d)
	{
		return std::chrono::duration<double, std::chrono::days::period>(d - epoch).count();
	}
    constexpr double to_excel(const ymd& d)
    {
        return to_excel(std::chrono::sys_days{d});
    }
	constexpr std::chrono::sys_days to_days(double d)
	{
		return epoch + std::chrono::days{ static_cast<std::chrono::sys_days::rep>(d) };
	}
	constexpr ymd to_ymd(double d)
	{
		return ymd{ to_days(d) };
	}
#ifdef _DEBUG
    static_assert(to_excel(epoch) == 0.0);
//	static_assert(to_excel(1900y/ 1 / 0) == 0.0);
#endif // _DEBUG

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