#include "dice_r_module.h"
#include <regex>
#include <string>
#include "cqsdk/cqsdk.h"
#include "dice_calculator.h"
#include "dice_exception.h"
#include "dice_utils.h"

namespace cq::event {
    struct MessageEvent;
}

namespace dice {
    bool r_module::match(const cq::event::MessageEvent& e, const std::wstring& ws) {
        std::wregex re(L"[ ]*[\\.。．][ ]*r.*", std::regex_constants::ECMAScript | std::regex_constants::icase);
        return std::regex_match(ws, re);
    }
    void r_module::process(const cq::event::MessageEvent& e, const std::wstring& ws) {
        std::wregex re(L"[ ]*[\\.。．][ ]*r(h)?[ ]*([0-9dk+\\-*/\\(\\)\\^bp]*)[ ]*(.*)",
                       std::regex_constants::ECMAScript | std::regex_constants::icase);
        std::wsmatch m;
        auto ret = std::regex_match(ws, m, re);
        std::wstring res;

        std::wstring dice(m[2]);
        std::wstring reason(m[3]);

        // 如果骰子全是数字，那就当成原因
        if (dice.find_first_not_of(L"0123456789") == std::wstring::npos) {
            reason = dice + reason;
            dice = L"";
        }

        if (!dice.empty()) {
            res = dice_calculator(dice, utils::get_defaultdice(e.target)).form_string();
        } else {
            res = dice_calculator(L"d", utils::get_defaultdice(e.target)).form_string();
        }

        if (m[1].first == m[1].second) {
            cq::api::send_msg(
                e.target,
                utils::format_string(msg::GetGlobalMsg("strRollDice"),
                                     std::map<std::string, std::string>{{"nick", utils::get_nickname(e.target)},
                                                                        {"reason", cq::utils::ws2s(reason)},
                                                                        {"dice_expression", cq::utils::ws2s(res)}}));
        } else {
            cq::api::send_msg(
                e.target,
                utils::format_string(msg::GetGlobalMsg("strHiddenDice"),
                                     std::map<std::string, std::string>{{"nick", utils::get_nickname(e.target)}}));
            cq::api::send_private_msg(
                *e.target.user_id,
                utils::format_string(msg::GetGlobalMsg("strRollHiddenDice"),
                                     std::map<std::string, std::string>{{"origin", utils::get_originname(e.target)},
                                                                        {"nick", utils::get_nickname(e.target)},
                                                                        {"reason", cq::utils::ws2s(reason)},
                                                                        {"dice_expression", cq::utils::ws2s(res)}}));
        }
    }
} // namespace dice